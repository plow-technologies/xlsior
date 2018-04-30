{-# LANGUAGE OverloadedStrings, FlexibleInstances, UndecidableInstances, FlexibleContexts #-}
module Xlsx.Parse where

import Data.XML.Types
import Text.XML
import Text.XML.Stream.Parse
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Vector as V
import Data.Vector (Vector)
import Data.ByteString (ByteString)
import qualified Data.ByteString.Lazy as L
import Data.Conduit
import qualified Data.Conduit.List as CL
import Control.Applicative hiding (many)
import Control.Monad.Reader
import Control.Monad.State
import Control.Monad.Trans.Resource
import Data.Attoparsec.Text
import Data.Default
import Data.Char
import Xlsx.Types

parseCoord :: Text -> Maybe (Int, Int)
parseCoord t = go 0 $ map ord $ T.unpack t where
    a = ord 'A'
    z = ord 'Z'
    zero = ord '0'
    nine = ord '9'
    go n [] = Nothing
    go n cs@(c:cs') | c <= z && c >= a = go (n*26 + c - a + 1) cs'
                    | n == 0 = Nothing
                    | otherwise = go' n 0 cs
    go' m n [] = Just (n, m)
    go' m n (c:cs) | c <= nine && c >= zero = go' m (n*10 + c - zero) cs
                   | otherwise = Nothing

tagLocal n = tag' $ matching (\Name {nameLocalName = l} -> n == l)

tagLocalNoAttr n c = tagIgnoreAttrs (matching (\Name {nameLocalName = l} -> n == l)) c

sharedStringSink :: (Monad m, MonadThrow m) => Sink ByteString m (Vector Text)
sharedStringSink = parseBytes def =$= (force "sst" $ tagLocal "sst" parseCount vecsink) where
    parseCount = requireAttr "uniqueCount" <* ignoreAttrs
    text = tagLocalNoAttr "t" $ content
    rich = do
        ts <- many $ tagLocalNoAttr "r" $ do
            tagLocalNoAttr "rPr" skipPr
            force "t" $ tagLocalNoAttr "t" $ content
        case ts of
            [] -> return Nothing
            l -> return $ Just $ T.concat l
    vecsink c = V.replicateM (read $ T.unpack c) $ force "si" $ tagLocalNoAttr "si" $
        force "r or t" $ orE text rich
    skipPr = do
        mbevent <- CL.peek
        case mbevent of
            Just j@(EventEndElement n) | nameLocalName n == "rPr" -> return ()
            Nothing -> throwM $ XmlException "no rPr closing tag" Nothing
            _ -> CL.drop 1 >> skipPr


rawRows :: (Monad m, MonadThrow m) => ConduitM ByteString [((Int, Int), Cell)] m ()
rawRows = parseBytes def =$= skiptorows where
    skiptorows = do
        mbevent <- await
        case mbevent of
            Nothing -> throwM $ XmlException "sheetData expected" Nothing
            Just j@(EventBeginElement n _) | nameLocalName n == "sheetData" -> rowsink
            _ -> skiptorows
    row = tagLocalNoAttr "row" $ many $ tagLocal "c" ((,) <$> requireAttr "r" <*> optionalAttr "t" <* ignoreAttrs) $ \(coord, mbt) -> do
        coord' <- case parseCoord coord of
            Nothing -> throwM $ XmlException ("invalid coordinate: " ++ T.unpack coord) Nothing
            Just j -> return j
        f <- tagLocalNoAttr "f" content
        cellv <- case mbt of
            Just typ  | typ == "inlineStr" -> InlineString <$> (force "is" $ tagLocalNoAttr "is" $ force "t" $ tagLocalNoAttr "t" content)
                      | otherwise -> do
                            v <- tagLocalNoAttr "v" content
                            return $ case v of
                                Nothing -> Empty
                                Just j | typ == "str" -> InlineString j
                                       | typ == "b" -> Boolean $ j == "1"
                                       | typ == "s" -> SharedString $ read $ T.unpack j
                                       | typ == "e" -> Error j
                                       | otherwise -> error $ show typ
            Nothing -> do
                v <- tagLocalNoAttr "v" content
                return $ case v of
                    Just j -> case parseOnly scientific j of
                        Left l -> error $ l
                        Right r -> Number r
                    _ -> Empty
        return (coord', (cellv, f))
    rowsink = do
        mbrow <- row
        case mbrow of
            Just j -> yield j >> rowsink
            Nothing -> return ()

unsparseSheet :: (Monad m, MonadThrow m) => ConduitM [((Int, Int), Cell)] [((Int, Int), Cell)] m ()
unsparseSheet = flip evalStateT 1 go where
    go = do
        mb <- lift await
        case mb of
            Nothing -> return ()
            Just r -> case r of
                [] -> lift $ yield []
                cs@(((r,_),_):_) -> do
                    cur <- get
                    if r > cur then lift $ replicateM_ (r - cur) (yield []) else return ()
                    lift $ yield cs
                    put $ r + 1
                    go

sheetRows :: (Applicative m, Monad m, MonadThrow m, FromRow a) => ConduitM ByteString a (ReaderT (Vector Text) m) ()
sheetRows = rawRows =$= unsparseSheet =$= CL.mapM fromRow
