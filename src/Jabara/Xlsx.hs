{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}
module Jabara.Xlsx (
  SheetName

  , readBook
  , readBookUnsafe

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

  , module Jabara.Xlsx.Types
) where

import           Codec.Xlsx
import           Control.Exception
import           Control.Lens
import           Data.ByteString.Lazy (readFile)
import           Data.Maybe
import           Data.Text
import           Data.Time.Calendar
import           Jabara.Xlsx.Types
import           Prelude              hiding (readFile)

readBook :: FilePath -> IO (Maybe Xlsx)
readBook path = (Just . toXlsx <$> readFile path) `catch` \(_::SomeException) -> pure Nothing

readBookUnsafe :: FilePath -> IO Xlsx
readBookUnsafe path = fromJust <$> readBook path

type SheetName = Text

cellBoolValueFromSheet :: Worksheet -> CellIndex -> Bool
cellBoolValueFromSheet sheet cell = cellBoolValue $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellBoolValueFromSheetM :: Worksheet -> CellIndex -> Maybe Bool
cellBoolValueFromSheetM sheet cell = cellBoolValueM $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellBoolValue :: Maybe CellValue -> Bool
cellBoolValue Nothing = error "cell not found"
cellBoolValue (Just (CellBool b)) = b
cellBoolValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellBoolValueM :: Maybe CellValue -> Maybe Bool
cellBoolValueM Nothing             = Nothing
cellBoolValueM (Just (CellBool b)) = Just b
cellBoolValueM (Just _)            = Nothing

cellDoubleValueFromSheet :: Worksheet -> CellIndex -> Double
cellDoubleValueFromSheet sheet cell = cellDoubleValue $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellDoubleValueFromSheetM :: Worksheet -> CellIndex -> Maybe Double
cellDoubleValueFromSheetM sheet cell = cellDoubleValueM $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellDoubleValue :: Maybe CellValue -> Double
cellDoubleValue Nothing = 0
cellDoubleValue (Just (CellDouble d)) = d
cellDoubleValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellDoubleValueM :: Maybe CellValue -> Maybe Double
cellDoubleValueM Nothing               = Nothing
cellDoubleValueM (Just (CellDouble d)) = Just d
cellDoubleValueM (Just _)              = Nothing

cellIntegralValueFromSheet :: (Integral a) => Worksheet -> CellIndex -> a
cellIntegralValueFromSheet sheet cell = cellIntegralValue $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellIntegralValueFromSheetM :: (Integral a) => Worksheet -> CellIndex -> Maybe a
cellIntegralValueFromSheetM sheet cell = cellIntegralValueM $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellIntegralValue :: (Integral a) => Maybe CellValue -> a
cellIntegralValue Nothing = 0
cellIntegralValue (Just (CellDouble d)) = ceiling d
cellIntegralValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellIntegralValueM :: (Integral a) => Maybe CellValue -> Maybe a
cellIntegralValueM Nothing               = Nothing
cellIntegralValueM (Just (CellDouble d)) = Just $ floor d
cellIntegralValueM (Just c)              = Nothing

cellStringValueFromSheet :: Worksheet -> CellIndex -> Text
cellStringValueFromSheet sheet cell = cellStringValue $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellStringValueFromSheetM :: Worksheet -> CellIndex -> Maybe Text
cellStringValueFromSheetM sheet cell = cellStringValueM $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellStringValue :: Maybe CellValue -> Text
cellStringValue Nothing = ""
cellStringValue (Just (CellText t)) = t
cellStringValue (Just c)            = error ("unexpected cell value --> " ++ show c)

cellStringValueM :: Maybe CellValue -> Maybe Text
cellStringValueM Nothing             = Nothing
cellStringValueM (Just (CellText t)) = Just t
cellStringValueM (Just _)            = Nothing

cellDayValueFromSheet :: Worksheet -> CellIndex -> Day
cellDayValueFromSheet sheet cell = cellDayValue $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellDayValueFromSheetM :: Worksheet -> CellIndex -> Maybe Day
cellDayValueFromSheetM sheet cell = cellDayValueM $ sheet ^? ixCell (toTuple cell) . cellValue . _Just

cellDayValue :: Maybe CellValue -> Day
cellDayValue Nothing = error "cell not found"
cellDayValue mCell = let dayValue = cellIntegralValue mCell
                     in  addDays (dayValue - 2) $ fromGregorian 1900 1 1

cellDayValueM :: Maybe CellValue -> Maybe Day
cellDayValueM Nothing = Nothing
cellDayValueM mCell = let dayValue = cellIntegralValue mCell
                      in  Just $ addDays (dayValue - 2) $ fromGregorian 1900 1 1
