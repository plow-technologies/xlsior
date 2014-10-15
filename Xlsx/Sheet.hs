{-# LANGUAGE OverloadedStrings, FlexibleInstances #-}
module Xlsx.Sheet where

import Data.Monoid
import Data.Char
import Text.Blaze
import Text.Blaze.Internal
import Data.ByteString (ByteString)
import Data.Text (Text)
import qualified Data.Text as T

type Row = Int -> Markup

{-newtype Rows = Rows { unRows :: Int -(Int, Markup) }

instance Monoid Rows where
    mappend a b = Rows $ \n -> Append (unRows a n) (unRows b n')
    mempty = undefined -}

class ToRow a where
    toRow :: a -> Row

decl :: Markup
decl = Content $ Static "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"

renderSheet :: [Row] -> Markup
renderSheet rows = Append decl $ wrksh where
    wrksh = AddAttribute "bla" " xmlns=\"" "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $
        AddAttribute "xmlns:r" " xmlns:r=\"" "http://schemas.openxmlformats.org/officeDocument/2006/relationships" $
        AddAttribute "xmlns:mc" " xmlns:mc=\"" "http://schemas.openxmlformats.org/markup-compatibility/2006" $
        AddAttribute "mc:Ignorable" " mc:Ignorable=\"" "x14ac" $
        AddAttribute "xmlns:x14ac" " xmlns:x14ac=\"" "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" $
        Parent "worksheet" "<worksheet" "</worksheet>" $
        Parent "sheetData" "<sheetData" "</sheetData>" $ go 1 rows
    go _ [] = Empty
    go n (r:rs) = Append (r n) (go (n+1) rs)

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

instance ToCell a => ToRow [a] where
    toRow cs = mkRow (spanRange $ length cs) $ \r -> mconcat $ map (\(n,c) -> cellName r n $ toCell c) $ zip [1..] cs

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

instance ToCell [Char] where
    toCell = inlineStr . toMarkup
