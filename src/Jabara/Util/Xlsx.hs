{-# LANGUAGE OverloadedStrings #-}
module Jabara.Util.Xlsx (
  RowIndex
  , ColumnIndex
  , SheetName
  , readBook

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
) where

import           Codec.Xlsx
import           Control.Lens
import           Data.ByteString.Lazy (readFile)
import           Data.Maybe
import           Data.Text
import           Data.Time.Calendar
import           Prelude              hiding (readFile)

readBook :: FilePath -> IO Xlsx
readBook path = toXlsx <$> readFile path

type SheetName = Text
type RowIndex = Int
type ColumnIndex = Int

cellBoolValueFromSheet :: Worksheet -> (RowIndex, ColumnIndex) -> Bool
cellBoolValueFromSheet sheet cell = cellBoolValue $ sheet ^? ixCell cell . cellValue . _Just

cellBoolValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Bool
cellBoolValueFromSheetM sheet cell = cellBoolValueM $ sheet ^? ixCell cell . cellValue . _Just

cellBoolValue :: Maybe CellValue -> Bool
cellBoolValue Nothing = error "cell not found"
cellBoolValue (Just (CellBool b)) = b
cellBoolValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellBoolValueM :: Maybe CellValue -> Maybe Bool
cellBoolValueM Nothing             = Nothing
cellBoolValueM (Just (CellBool b)) = Just b
cellBoolValueM (Just _)            = Nothing

cellDoubleValueFromSheet :: Worksheet -> (RowIndex, ColumnIndex) -> Double
cellDoubleValueFromSheet sheet cell = cellDoubleValue $ sheet ^? ixCell cell . cellValue . _Just

cellDoubleValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Double
cellDoubleValueFromSheetM sheet cell = cellDoubleValueM $ sheet ^? ixCell cell . cellValue . _Just

cellDoubleValue :: Maybe CellValue -> Double
cellDoubleValue Nothing = 0
cellDoubleValue (Just (CellDouble d)) = d
cellDoubleValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellDoubleValueM :: Maybe CellValue -> Maybe Double
cellDoubleValueM Nothing               = Nothing
cellDoubleValueM (Just (CellDouble d)) = Just d
cellDoubleValueM (Just _)              = Nothing

cellIntegralValueFromSheet :: (Integral a) => Worksheet -> (RowIndex, ColumnIndex) -> a
cellIntegralValueFromSheet sheet cell = cellIntegralValue $ sheet ^? ixCell cell . cellValue . _Just

cellIntegralValueFromSheetM :: (Integral a) => Worksheet -> (RowIndex, ColumnIndex) -> Maybe a
cellIntegralValueFromSheetM sheet cell = cellIntegralValueM $ sheet ^? ixCell cell . cellValue . _Just

cellIntegralValue :: (Integral a) => Maybe CellValue -> a
cellIntegralValue Nothing = 0
cellIntegralValue (Just (CellDouble d)) = ceiling d
cellIntegralValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellIntegralValueM :: (Integral a) => Maybe CellValue -> Maybe a
cellIntegralValueM Nothing               = Nothing
cellIntegralValueM (Just (CellDouble d)) = Just $ floor d
cellIntegralValueM (Just c)              = Nothing

cellStringValueFromSheet :: Worksheet -> (RowIndex, ColumnIndex) -> Text
cellStringValueFromSheet sheet cell = cellStringValue $ sheet ^? ixCell cell . cellValue . _Just

cellStringValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Text
cellStringValueFromSheetM sheet cell = cellStringValueM $ sheet ^? ixCell cell . cellValue . _Just

cellStringValue :: Maybe CellValue -> Text
cellStringValue Nothing = ""
cellStringValue (Just (CellText t)) = t
cellStringValue (Just c)            = error ("unexpected cell value --> " ++ show c)

cellStringValueM :: Maybe CellValue -> Maybe Text
cellStringValueM Nothing             = Nothing
cellStringValueM (Just (CellText t)) = Just t
cellStringValueM (Just _)            = Nothing

cellDayValueFromSheet :: Worksheet -> (RowIndex, ColumnIndex) -> Day
cellDayValueFromSheet sheet cell = cellDayValue $ sheet ^? ixCell cell . cellValue . _Just

cellDayValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Day
cellDayValueFromSheetM sheet cell = cellDayValueM $ sheet ^? ixCell cell . cellValue . _Just

cellDayValue :: Maybe CellValue -> Day
cellDayValue Nothing = error "cell not found"
cellDayValue mCell = let dayValue = cellIntegralValue mCell
                     in  addDays (dayValue - 2) $ fromGregorian 1900 1 1

cellDayValueM :: Maybe CellValue -> Maybe Day
cellDayValueM Nothing = Nothing
cellDayValueM mCell = let dayValue = cellIntegralValue mCell
                      in  Just $ addDays (dayValue - 2) $ fromGregorian 1900 1 1
