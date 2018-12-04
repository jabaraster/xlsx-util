{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}
module Jabara.Util.Xlsx (
  RowIndex
  , ColumnIndex
  , SheetName
  , CellLocation(..)

  , readBook
  , readBookUnsafe

  , parseCellLocationText

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
import           Control.Exception
import           Control.Lens
import           Data.ByteString.Lazy (readFile)
import           Data.Maybe
import           Data.Text
import           Data.Time.Calendar
import           Prelude              hiding (readFile)
import qualified Text.Parsec as TP

readBook :: FilePath -> IO (Maybe Xlsx)
readBook path = (Just . toXlsx <$> readFile path) `catch` \(_::SomeException) -> pure Nothing

readBookUnsafe :: FilePath -> IO Xlsx
readBookUnsafe path = fromJust <$> readBook path

type SheetName = Text
type RowIndex = Int
type ColumnIndex = Int

newtype CellLocation = CellLocation (RowIndex, ColumnIndex)
  deriving (Show, Eq, Read)

cellBoolValueFromSheet :: Worksheet -> CellLocation -> Bool
cellBoolValueFromSheet sheet (CellLocation (r,c)) = cellBoolValue $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellBoolValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Bool
cellBoolValueFromSheetM sheet (r,c) = cellBoolValueM $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellBoolValue :: Maybe CellValue -> Bool
cellBoolValue Nothing = error "cell not found"
cellBoolValue (Just (CellBool b)) = b
cellBoolValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellBoolValueM :: Maybe CellValue -> Maybe Bool
cellBoolValueM Nothing             = Nothing
cellBoolValueM (Just (CellBool b)) = Just b
cellBoolValueM (Just _)            = Nothing

cellDoubleValueFromSheet :: Worksheet -> (RowIndex, ColumnIndex) -> Double
cellDoubleValueFromSheet sheet (r,c) = cellDoubleValue $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellDoubleValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Double
cellDoubleValueFromSheetM sheet (r,c) = cellDoubleValueM $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellDoubleValue :: Maybe CellValue -> Double
cellDoubleValue Nothing = 0
cellDoubleValue (Just (CellDouble d)) = d
cellDoubleValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellDoubleValueM :: Maybe CellValue -> Maybe Double
cellDoubleValueM Nothing               = Nothing
cellDoubleValueM (Just (CellDouble d)) = Just d
cellDoubleValueM (Just _)              = Nothing

cellIntegralValueFromSheet :: (Integral a) => Worksheet -> (RowIndex, ColumnIndex) -> a
cellIntegralValueFromSheet sheet (r,c) = cellIntegralValue $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellIntegralValueFromSheetM :: (Integral a) => Worksheet -> (RowIndex, ColumnIndex) -> Maybe a
cellIntegralValueFromSheetM sheet (r,c) = cellIntegralValueM $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellIntegralValue :: (Integral a) => Maybe CellValue -> a
cellIntegralValue Nothing = 0
cellIntegralValue (Just (CellDouble d)) = ceiling d
cellIntegralValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellIntegralValueM :: (Integral a) => Maybe CellValue -> Maybe a
cellIntegralValueM Nothing               = Nothing
cellIntegralValueM (Just (CellDouble d)) = Just $ floor d
cellIntegralValueM (Just c)              = Nothing

cellStringValueFromSheet :: Worksheet -> CellLocation -> Text
cellStringValueFromSheet sheet (CellLocation (r,c)) = cellStringValue $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellStringValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Text
cellStringValueFromSheetM sheet (r,c) = cellStringValueM $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellStringValue :: Maybe CellValue -> Text
cellStringValue Nothing = ""
cellStringValue (Just (CellText t)) = t
cellStringValue (Just c)            = error ("unexpected cell value --> " ++ show c)

cellStringValueM :: Maybe CellValue -> Maybe Text
cellStringValueM Nothing             = Nothing
cellStringValueM (Just (CellText t)) = Just t
cellStringValueM (Just _)            = Nothing

cellDayValueFromSheet :: Worksheet -> (RowIndex, ColumnIndex) -> Day
cellDayValueFromSheet sheet (r,c) = cellDayValue $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellDayValueFromSheetM :: Worksheet -> (RowIndex, ColumnIndex) -> Maybe Day
cellDayValueFromSheetM sheet (r,c) = cellDayValueM $ sheet ^? ixCell (r+1,c+1) . cellValue . _Just

cellDayValue :: Maybe CellValue -> Day
cellDayValue Nothing = error "cell not found"
cellDayValue mCell = let dayValue = cellIntegralValue mCell
                     in  addDays (dayValue - 2) $ fromGregorian 1900 1 1

cellDayValueM :: Maybe CellValue -> Maybe Day
cellDayValueM Nothing = Nothing
cellDayValueM mCell = let dayValue = cellIntegralValue mCell
                      in  Just $ addDays (dayValue - 2) $ fromGregorian 1900 1 1

parseCellLocationText :: Text -> Maybe CellLocation
parseCellLocationText s = case parseCore s of
  Left _ -> Nothing
  Right cell -> Just cell
  where
    parseCore :: Text -> Either TP.ParseError CellLocation
    parseCore = TP.parse coreParser "" . unpack

    coreParser :: TP.ParsecT String u Identity CellLocation
    coreParser = do
      colString <- TP.many1 TP.upper
      row <- rowParser
      return $ CellLocation (row, 0)

    rowParser :: TP.ParsecT String u Identity ColumnIndex
    rowParser = do
      rowString <- TP.many1 TP.digit
      return $ flip (-) 1 $ read rowString
