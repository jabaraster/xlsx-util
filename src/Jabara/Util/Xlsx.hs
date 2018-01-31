{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RankNTypes        #-}
{-# LANGUAGE TemplateHaskell   #-}
module Jabara.Util.Xlsx (
    RowIndex
  , ColumnIndex
  , SheetName
  , readBook

  , CellLocation(..)
  , clRow
  , clColumn
  , shiftCellLocation
  , shiftRowLocation
  , shiftColumnLocation

  , cellBoolValue
  , cellBoolValueFromSheet
  , cellBoolValueM
  , cellBoolValueFromSheetM

  , cellDoubleValue
  , cellDoubleValueFromSheet
  , cellDoubleValueM
  , cellDoubleValueFromSheetM

  , cellIntegralValue
  , cellIntegralValueFromSheet
  , cellIntegralValueM
  , cellIntegralValueFromSheetM

  , cellStringValue
  , cellStringValueFromSheet
  , cellStringValueM
  , cellStringValueFromSheetM

  , cellDayValue
  , cellDayValueFromSheet
  , cellDayValueM
  , cellDayValueFromSheetM

  , columnLabelToIndex
  , parseCellName
) where

import           Codec.Xlsx
import           Control.Lens
import           Data.ByteString.Lazy (readFile)
import           Data.Maybe
import           Data.Text            (Text, pack, toUpper, unpack)
import           Data.Time.Calendar
import           Prelude              hiding (readFile)
import           Text.Parsec

readBook :: FilePath -> IO Xlsx
readBook path = readFile path >>= pure . toXlsx

type SheetName = Text
type RowIndex = Int
type ColumnIndex = Int
type ColumnLabel = Text

data CellLocation =
    CellLocation {
      _clRow    :: Int
    , _clColumn :: Int
    } deriving (Show, Read, Eq)
makeLenses ''CellLocation

shiftCellLocation :: CellLocation -> CellLocation -> CellLocation
shiftCellLocation base offset = CellLocation {
                              _clRow    = base^.clRow + offset^.clRow
                            , _clColumn = base^.clColumn + offset^.clColumn
                            }

shiftRowLocation :: CellLocation -> Int -> CellLocation
shiftRowLocation base offset = base { _clRow = base^.clRow + offset }

shiftColumnLocation :: CellLocation -> Int -> CellLocation
shiftColumnLocation base offset = base { _clColumn = base^.clColumn + offset }

columnLabelToIndex :: ColumnLabel -> ColumnIndex
columnLabelToIndex label = fromInteger $ snd $ foldr core (0, 0)  $ unpack $ toUpper label
  where
    core :: Char -> (Integer, Integer) -> (Integer, Integer)
    core alphabet (digit, value) =
        let val = (26 ^ digit) * ((fromEnum alphabet) - 64)
        in  (digit + 1, value + toInteger val)

parseCellName :: Text -> Maybe CellLocation
parseCellName n =
    case runParser p () "" (unpack n) of
        Left  _ -> Nothing
        Right r -> Just r
  where
    p = do
          col <- many letter
          row <- many $ oneOf ['1','2','3','4','5','6','7','8','9','0']
          return CellLocation {
                   _clRow    = read row
                 , _clColumn = columnLabelToIndex $ pack col
                 }

ixCell' :: CellLocation -> Traversal' Worksheet Cell
ixCell' ci = ixCell (ci^.clRow, ci^.clColumn)

cellBoolValueFromSheet :: Worksheet -> CellLocation -> Bool
cellBoolValueFromSheet sheet cell = cellBoolValue $ sheet ^? ixCell' cell . cellValue . _Just

cellBoolValueFromSheetM :: Worksheet -> CellLocation -> Maybe Bool
cellBoolValueFromSheetM sheet cell = cellBoolValueM $ sheet ^? ixCell' cell . cellValue . _Just

cellBoolValue :: Maybe CellValue -> Bool
cellBoolValue Nothing = error "cell not found"
cellBoolValue (Just (CellBool b)) = b
cellBoolValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellBoolValueM :: Maybe CellValue -> Maybe Bool
cellBoolValueM Nothing             = Nothing
cellBoolValueM (Just (CellBool b)) = Just b
cellBoolValueM (Just _)            = Nothing

cellDoubleValueFromSheet :: Worksheet -> CellLocation -> Double
cellDoubleValueFromSheet sheet cell = cellDoubleValue $ sheet ^? ixCell' cell . cellValue . _Just

cellDoubleValueFromSheetM :: Worksheet -> CellLocation -> Maybe Double
cellDoubleValueFromSheetM sheet cell = cellDoubleValueM $ sheet ^? ixCell' cell . cellValue . _Just

cellDoubleValue :: Maybe CellValue -> Double
cellDoubleValue Nothing = 0
cellDoubleValue (Just (CellDouble d)) = d
cellDoubleValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellDoubleValueM :: Maybe CellValue -> Maybe Double
cellDoubleValueM Nothing               = Nothing
cellDoubleValueM (Just (CellDouble d)) = Just d
cellDoubleValueM (Just _)              = Nothing

cellIntegralValueFromSheet :: (Integral a) => Worksheet -> CellLocation -> a
cellIntegralValueFromSheet sheet cell = cellIntegralValue $ sheet ^? ixCell' cell . cellValue . _Just

cellIntegralValueFromSheetM :: (Integral a) => Worksheet -> CellLocation -> Maybe a
cellIntegralValueFromSheetM sheet cell = cellIntegralValueM $ sheet ^? ixCell' cell . cellValue . _Just

cellIntegralValue :: (Integral a) => Maybe CellValue -> a
cellIntegralValue Nothing = 0
cellIntegralValue (Just (CellDouble d)) = ceiling d
cellIntegralValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellIntegralValueM :: (Integral a) => Maybe CellValue -> Maybe a
cellIntegralValueM Nothing               = Nothing
cellIntegralValueM (Just (CellDouble d)) = Just $ floor d
cellIntegralValueM (Just c)              = Nothing

cellStringValueFromSheet :: Worksheet -> CellLocation -> Text
cellStringValueFromSheet sheet cell = cellStringValue $ sheet ^? ixCell' cell . cellValue . _Just

cellStringValueFromSheetM :: Worksheet -> CellLocation -> Maybe Text
cellStringValueFromSheetM sheet cell = cellStringValueM $ sheet ^? ixCell' cell . cellValue . _Just

cellStringValue :: Maybe CellValue -> Text
cellStringValue Nothing = ""
cellStringValue (Just (CellText t)) = t
cellStringValue (Just c)            = error ("unexpected cell value --> " ++ show c)

cellStringValueM :: Maybe CellValue -> Maybe Text
cellStringValueM Nothing             = Nothing
cellStringValueM (Just (CellText t)) = Just t
cellStringValueM (Just _)            = Nothing

cellDayValueFromSheet :: Worksheet -> CellLocation -> Day
cellDayValueFromSheet sheet cell = cellDayValue $ sheet ^? ixCell' cell . cellValue . _Just

cellDayValueFromSheetM :: Worksheet -> CellLocation -> Maybe Day
cellDayValueFromSheetM sheet cell = cellDayValueM $ sheet ^? ixCell' cell . cellValue . _Just

cellDayValue :: Maybe CellValue -> Day
cellDayValue Nothing = error "cell not found"
cellDayValue mCell = let dayValue = cellIntegralValue mCell
                     in  addDays (dayValue - 2) $ fromGregorian 1900 1 1

cellDayValueM :: Maybe CellValue -> Maybe Day
cellDayValueM Nothing = Nothing
cellDayValueM mCell = let dayValue = cellIntegralValue mCell
                      in  Just $ addDays (dayValue - 2) $ fromGregorian 1900 1 1
