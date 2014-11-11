module Xlsx.Types.Internal where

import Text.Blaze.Internal
import Data.Text (Text)

data CellValue = InlineString Text | SharedString Int | Number Double | Boolean Bool | Date Text | Error Text | Empty deriving (Read, Show, Eq)
type Cell = (CellValue, Maybe Text)

type Row = Int -> Markup
