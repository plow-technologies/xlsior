{-# LANGUAGE OverloadedStrings, FlexibleInstances #-}
module Xlsx.Sheet where

import Data.Monoid
import Data.Char
import Text.Blaze
import Text.Blaze.Internal
import Data.ByteString (ByteString)
import Data.Text (Text)
import qualified Data.Text as T
import Xlsx.Types hiding (Empty)

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
