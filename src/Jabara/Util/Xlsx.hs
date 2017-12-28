{-# LANGUAGE OverloadedStrings #-}
module Jabara.Util.Xlsx (
  RowIndex
  , ColumnIndex
  , SheetName
  , readBook
  , cellBoolValue
  , cellBoolValue'
  , cellDoubleValue
  , cellDoubleValue'
  , cellIntegralValue
  , cellIntegralValue'
  , cellStringValue
  , cellStringValue'
  , cellDayValue
  , cellDayValue'
  , cellDayValueM
  , cellDayValueM'
) where

import           Codec.Xlsx
import           Control.Lens
import           Data.ByteString.Lazy (readFile)
import           Data.Maybe
import           Data.Text
import           Data.Time.Calendar
import           Prelude              hiding (readFile)

readBook :: FilePath -> IO Xlsx
readBook path = readFile path >>= pure . toXlsx

type SheetName = Text
type RowIndex = Int
type ColumnIndex = Int

cellBoolValue' :: Worksheet -> (RowIndex, ColumnIndex) -> Bool
cellBoolValue' sheet cell = cellBoolValue $ sheet ^? ixCell cell . cellValue . _Just

cellBoolValue :: Maybe CellValue -> Bool
cellBoolValue Nothing = error "cell not found"
cellBoolValue (Just (CellBool b)) = b
cellBoolValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellDoubleValue' :: Worksheet -> (RowIndex, ColumnIndex) -> Double
cellDoubleValue' sheet cell = cellDoubleValue $ sheet ^? ixCell cell . cellValue . _Just

cellDoubleValue :: Maybe CellValue -> Double
cellDoubleValue Nothing = 0
cellDoubleValue (Just (CellDouble d)) = d
cellDoubleValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellIntegralValue' :: (Integral a) => Worksheet -> (RowIndex, ColumnIndex) -> a
cellIntegralValue' sheet cell = cellIntegralValue $ sheet ^? ixCell cell . cellValue . _Just

cellIntegralValue :: (Integral a) => Maybe CellValue -> a
cellIntegralValue Nothing = 0
cellIntegralValue (Just (CellDouble d)) = ceiling d
cellIntegralValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellStringValue' :: Worksheet -> (RowIndex, ColumnIndex) -> Text
cellStringValue' sheet cell = cellStringValue $ sheet ^? ixCell cell . cellValue . _Just

cellStringValue :: Maybe CellValue -> Text
cellStringValue Nothing = ""
cellStringValue (Just (CellText t)) = t
cellStringValue (Just c)            = error ("unexpected cell value --> " ++ show c)


cellDayValue' :: Worksheet -> (RowIndex, ColumnIndex) -> Day
cellDayValue' sheet cell = cellDayValue $ sheet ^? ixCell cell . cellValue . _Just

cellDayValue :: Maybe CellValue -> Day
cellDayValue Nothing = error "cell not found"
cellDayValue mCell = let dayValue = cellIntegralValue mCell
                     in  addDays (dayValue - 2) $ fromGregorian 1900 1 1

cellDayValueM' :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Day
cellDayValueM' sheet cell = cellDayValueM $ sheet ^? ixCell cell . cellValue . _Just

cellDayValueM :: Maybe CellValue -> Maybe Day
cellDayValueM Nothing = Nothing
cellDayValueM mCell = let dayValue = cellIntegralValue mCell
                      in  Just $ addDays (dayValue - 2) $ fromGregorian 1900 1 1
