{-# LANGUAGE OverloadedStrings, FlexibleInstances, UndecidableInstances #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}
module Xlsx.Types.Instances where

import Control.Applicative
import Control.Monad.Reader
import Control.Monad.Identity
import Data.Char
import Data.Scientific
import Data.Monoid
import Data.Text (Text)
import qualified Data.Text as T
import Data.ByteString (ByteString)
import qualified Data.Vector as V
import Text.Blaze
import Text.Blaze.Internal hiding (Empty)
import Xlsx.Types.Internal
import Xlsx.Types.Class

instance ToCell Markup where
    toCell = id

instance ToCell Integer where
    toCell = numericCell

instance ToCell Int where
    toCell = numericCell

instance ToCell Float where
    toCell = numericCell

instance ToCell Double where
    toCell = numericCell

inlineStr :: Markup -> Markup
inlineStr s = AddAttribute "t" " t=\"" "inlineStr" $ Parent "c" "<c" "</c>" $ Parent "is" "<is" "</is>" $
    Parent "t" "<t" "</t>" s

instance ToCell ByteString where
    toCell = inlineStr . unsafeByteString

instance ToCell Text where
    toCell = inlineStr . toMarkup

instance ToCell [Char] where
    toCell = inlineStr . toMarkup

rowValue :: Int -> ChoiceString
rowValue n = Text $ T.pack $ show n

spanRange :: Int -> ChoiceString
spanRange n = Text $ "1:" <> (T.pack $ show n)

cellName :: Int -> Int -> Markup -> Markup
cellName r c = AddAttribute "r" " r=\"" cell where
    cell = Text $ T.pack $ int2col c <> show r
    int2col = reverse . map int2let . base26 where
        int2let 0 = 'Z'
        int2let x = chr $ (x - 1) + ord 'A'
        base26  0 = []
        base26  i = let i' = (i `mod` 26)
                        i'' = if i' == 0 then 26 else i'
                    in seq i' (i' : base26 ((i - i'') `div` 26))

mkRow :: ChoiceString -> (Int -> Markup) -> Row
mkRow spn m n = AddAttribute "r" " r=\"" (rowValue n) $ AddAttribute "spans" " spans=\"" spn $
    Parent "row" "<row" "</row>" $ m n

instance (ToCell a, ToCell b) => ToRow (a, b) where
    toRow (a, b) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) where
            spn = spanRange 2

instance (ToCell a, ToCell b, ToCell c) => ToRow (a, b, c) where
    toRow (a, b, c) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) where
            spn = spanRange 3

instance (ToCell a, ToCell b, ToCell c, ToCell d) => ToRow (a, b, c, d) where
    toRow (a, b, c, d) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) where
            spn = spanRange 3

instance (ToCell a, ToCell b, ToCell c, ToCell d, ToCell e) => ToRow (a, b, c, d, e) where
    toRow (a, b, c, d, e) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) <> (cellName r 5 $ toCell e) where
            spn = spanRange 5

instance (ToCell a, ToCell b, ToCell c, ToCell d, ToCell e, ToCell f) => ToRow (a, b, c, d, e, f) where
    toRow (a, b, c, d, e, f) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) <> (cellName r 5 $ toCell e) <> (cellName r 6 $ toCell f) where
            spn = spanRange 6

instance (ToCell a, ToCell b, ToCell c, ToCell d, ToCell e, ToCell f, ToCell g) => ToRow (a, b, c, d, e, f, g) where
    toRow (a, b, c, d, e, f, g) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) <> (cellName r 5 $ toCell e) <> (cellName r 6 $ toCell f) <>
        (cellName r 7 $ toCell g) where
            spn = spanRange 7

instance (ToCell a, ToCell b, ToCell c, ToCell d, ToCell e, ToCell f, ToCell g, ToCell h) =>
    ToRow (a, b, c, d, e, f, g, h) where
    toRow (a, b, c, d, e, f, g, h) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) <> (cellName r 5 $ toCell e) <> (cellName r 6 $ toCell f) <>
        (cellName r 7 $ toCell g) <> (cellName r 8 $ toCell h) where
            spn = spanRange 8

instance (ToCell a, ToCell b, ToCell c, ToCell d, ToCell e, ToCell f, ToCell g, ToCell h, ToCell i) =>
    ToRow (a, b, c, d, e, f, g, h, i) where
    toRow (a, b, c, d, e, f, g, h, i) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) <> (cellName r 5 $ toCell e) <> (cellName r 6 $ toCell f) <>
        (cellName r 7 $ toCell g) <> (cellName r 8 $ toCell h) <> (cellName r 9 $ toCell i)where
            spn = spanRange 9

instance ToCell a => ToRow [a] where
    toRow cs = mkRow (spanRange $ length cs) $ \r -> mconcat $ map (\(n,c) -> cellName r n $ toCell c) $ zip [1..] cs

numericCell :: (Num a, ToMarkup a) => a -> Markup
numericCell v = Parent "c" "<c" "</c>" $ Parent "v" "<v" "</v>" $ toMarkup v

unsparse :: [((Int, Int), Cell)] -> [Cell]
unsparse cols = go 1 cols where
    go _ [] = []
    go n cs@(((_,n'),c):cs') | n' > n = (Empty, Nothing) : go (n+1) cs
                             | otherwise = c : go (n+1) cs'

instance FromCell (Text) where
    fromCell (SharedString si,_) = (V.! si) <$> ask
    fromCell (InlineString s,_) = pure s
    fromCell _ = error "expected text"

instance FromCell (Int) where
    fromCell (Number n,_) = pure $ floor n
    fromCell _ = error "expected number"

instance FromCell (Double) where
    fromCell (Number n,_) = pure $ toRealFloat n
    fromCell _ = error "expected number"

instance FromCell (Bool) where
    fromCell (Number n,_) = pure $ n == 1
    fromCell _ = error "expected bool"

instance (FromCell a) => FromCell (Maybe a) where
    fromCell (Error {},_) = pure Nothing
    fromCell (Empty,_) = pure Nothing
    fromCell o = Just <$> fromCell o

instance (FromCell a) => FromRow (Identity a) where
    fromRow r = case unsparse r of
        [a] -> Identity <$> fromCell a
        _ -> error "columns"

instance (FromCell a, FromCell b) => FromRow (a,b) where
    fromRow r = case unsparse r of
        [a,b] -> (,) <$> fromCell a <*> fromCell b
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c) => FromRow (a,b,c) where
    fromRow r = case unsparse r of
        [a,b,c] -> (,,) <$> fromCell a <*> fromCell b <*> fromCell c
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d) => FromRow (a,b,c,d) where
    fromRow r = case unsparse r of
        [a,b,c,d] -> (,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e) => FromRow (a,b,c,d,e) where
    fromRow r = case unsparse r of
        [a,b,c,d,e] -> (,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f) => FromRow (a,b,c,d,e,f) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f] -> (,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g) => FromRow (a,b,c,d,e,f,g) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g] -> (,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h) => FromRow (a,b,c,d,e,f,g,h) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h] -> (,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i) => FromRow (a,b,c,d,e,f,g,h,i) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i] -> (,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j) => FromRow (a,b,c,d,e,f,g,h,i,j) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j] -> (,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k) => FromRow (a,b,c,d,e,f,g,h,i,j,k) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k] -> (,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l] -> (,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m] -> (,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n] -> (,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o] -> (,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p] -> (,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q] -> (,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r] -> (,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s] -> (,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s, FromCell t) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t] -> (,,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s <*> fromCell t
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s, FromCell t, FromCell u) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u] -> (,,,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s <*> fromCell t <*> fromCell u
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s, FromCell t, FromCell u, FromCell v) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v] -> (,,,,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s <*> fromCell t <*> fromCell u <*> fromCell v
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s, FromCell t, FromCell u, FromCell v, FromCell w) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w] -> (,,,,,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s <*> fromCell t <*> fromCell u <*> fromCell v <*> fromCell w
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s, FromCell t, FromCell u, FromCell v, FromCell w, FromCell x) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x] -> (,,,,,,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s <*> fromCell t <*> fromCell u <*> fromCell v <*> fromCell w <*> fromCell x
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s, FromCell t, FromCell u, FromCell v, FromCell w, FromCell x, FromCell y) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y] -> (,,,,,,,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s <*> fromCell t <*> fromCell u <*> fromCell v <*> fromCell w <*> fromCell x <*> fromCell y
        _ -> error "columns"

instance (FromCell a, FromCell b, FromCell c, FromCell d, FromCell e, FromCell f, FromCell g, FromCell h, FromCell i, FromCell j, FromCell k, FromCell l, FromCell m, FromCell n, FromCell o, FromCell p, FromCell q, FromCell r, FromCell s, FromCell t, FromCell u, FromCell v, FromCell w, FromCell x, FromCell y, FromCell z) => FromRow (a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z) where
    fromRow r = case unsparse r of
        [a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z] -> (,,,,,,,,,,,,,,,,,,,,,,,,,) <$> fromCell a <*> fromCell b <*> fromCell c <*> fromCell d <*> fromCell e <*> fromCell f <*> fromCell g <*> fromCell h <*> fromCell i <*> fromCell j <*> fromCell k <*> fromCell l <*> fromCell m <*> fromCell n <*> fromCell o <*> fromCell p <*> fromCell q <*> fromCell r <*> fromCell s <*> fromCell t <*> fromCell u <*> fromCell v <*> fromCell w <*> fromCell x <*> fromCell y <*> fromCell z
        _ -> error "columns"
