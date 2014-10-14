{-# LANGUAGE OverloadedStrings #-}
module Xlsx.Types where

import Data.Monoid
import Data.Char
import Text.Blaze
import Text.Blaze.Internal
import Text.Blaze.Renderer.Utf8
import Data.ByteString (ByteString)
import Data.Text (Text)
import qualified Data.Text as T

newtype SheetData = SheetData { unRows :: Int -> Markup }

instance Monoid SheetData where
    mappend a b = SheetData $ \n -> Append (unRows a n) (unRows b (n+1))
    mempty = undefined

class ToRows a where
    toRows :: a -> SheetData

renderSheet :: SheetData -> Markup
renderSheet s = Append decl $ wrksh where
    decl = undefined
    wrksh = undefined

fromRows :: ToRows a => [a] -> SheetData
fromRows = mconcat . map toRows

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

mkRow :: ChoiceString -> (Int -> Markup) -> SheetData
mkRow spn m = SheetData $ \n -> AddAttribute "r" " r=\"" (rowValue n) $ AddAttribute "spans" " spans=\"" spn $
    Parent "row" "<row" "</row" $ m n

instance (ToCell a, ToCell b) => ToRows (a, b) where
    toRows (a, b) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) where
            spn = spanRange 2

instance (ToCell a, ToCell b, ToCell c) => ToRows (a, b, c) where
    toRows (a, b, c) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) where
            spn = spanRange 3

instance (ToCell a, ToCell b, ToCell c, ToCell d) => ToRows (a, b, c, d) where
    toRows (a, b, c, d) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) where
            spn = spanRange 3

instance (ToCell a, ToCell b, ToCell c, ToCell d, ToCell e) => ToRows (a, b, c, d, e) where
    toRows (a, b, c, d, e) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) <> (cellName r 5 $ toCell e) where
            spn = spanRange 5

instance (ToCell a, ToCell b, ToCell c, ToCell d, ToCell e, ToCell f) => ToRows (a, b, c, d, e, f) where
    toRows (a, b, c, d, e, f) = mkRow spn $ \r ->
        (cellName r 1 $ toCell a) <> (cellName r 2 $ toCell b) <> (cellName r 3 $ toCell c) <>
        (cellName r 4 $ toCell d) <> (cellName r 5 $ toCell e) <> (cellName r 6 $ toCell f) where
            spn = spanRange 6

class ToCell a where
    toCell :: a -> Markup

numericCell :: (Num a, ToMarkup a) => a -> Markup
numericCell v = Parent "c" "<c" "</c>" $ Parent "v" "<v" "</v>" $ toMarkup v

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
    
