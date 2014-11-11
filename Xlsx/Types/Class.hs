module Xlsx.Types.Class where

import Xlsx.Types.Internal
import Control.Applicative
import Control.Monad.Reader
import Data.Vector (Vector)
import Data.Text (Text)
import Text.Blaze.Internal

class FromRow a where
    fromRow :: (Applicative m, Monad m) => [((Int, Int), Cell)] -> ReaderT (Vector Text) m a

class FromCell a where
    fromCell :: (Applicative m, Monad m) => Cell -> ReaderT (Vector Text) m a

class ToRow a where
    toRow :: a -> Row

class ToCell a where
    toCell :: a -> Markup

