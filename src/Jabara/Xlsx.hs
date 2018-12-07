{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}
module Jabara.Xlsx (
  SheetName

  , readBook
  , readBookUnsafe

  , writeBook
  , writeBookUnsafe

  , cell
  , cellUnsafe
  , cellT
  , cellTUnsafe
  , columnCellT
  , columnCellTUnsafe
  , moveRow
  , moveColumn
  , tuple
  , fromTuple

  , (+++)
  , (^^^)
  , (>>>)
  , (+^^)
  , (+>>)

  , parseColumnIndexTextUnsafe

  , parseCellIndexTextUnsafe

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
import           Data.ByteString.Lazy  (readFile, writeFile)
import           Data.Either
import           Data.Maybe
import qualified Data.Text             as DT (Text, foldr)
import           Data.Time.Calendar
import           Data.Time.Clock.POSIX
import           Jabara.Xlsx.Types
import           Prelude               hiding (readFile, writeFile)

readBook :: FilePath -> IO (Either SomeException Xlsx)
readBook path = (Right . toXlsx <$> readFile path) `catch` (pure . Left)

readBookUnsafe :: FilePath -> IO Xlsx
readBookUnsafe path = do
  eBook <- readBook path
  case eBook of
    Left  ex   -> error $ displayException ex
    Right book -> pure book

writeBook :: FilePath -> Xlsx -> IO (Either SomeException ())
writeBook path book = do
  ct <- getPOSIXTime
  let bytes = fromXlsx ct book
  writeFile path bytes >> pure (Right ())
    `catch` (pure . Left)

writeBookUnsafe :: FilePath -> Xlsx -> IO ()
writeBookUnsafe path book = do
  e <- writeBook path book
  case e of
    Left e -> error $ displayException e
    Right _ -> pure ()

type SheetName = DT.Text

{- |
  Codec.Xlsモジュールの、ixCellなどセルを扱う際のタプルに変換する.
-}
tuple :: CellIndex -> (Int, Int)
tuple ci = (rowIndexValue (ciRowIndex ci) + 1, columnIndexValue (ciColumnIndex ci) + 1)

fromTuple :: (Int, Int) -> CellIndex
fromTuple (r, c) = CellIndex (RI (r - 1)) (CI (c- 1))

{- |
  文字列をixCellなどセルを扱う際のタプルに変換する.
-}
cellT :: DT.Text -> Maybe (Int, Int)
cellT s = tuple <$> parseCellIndexText s

cellTUnsafe :: DT.Text -> (Int, Int)
cellTUnsafe = tuple . parseCellIndexTextUnsafe

cell :: DT.Text -> Maybe CellIndex
cell = parseCellIndexText

cellUnsafe :: DT.Text -> CellIndex
cellUnsafe = parseCellIndexTextUnsafe

columnCellT :: RowIndex -> DT.Text -> Maybe CellIndexTuple
columnCellT row t = tuple . CellIndex row <$> parseColumnIndexText t

columnCellTUnsafe :: RowIndex -> DT.Text -> CellIndexTuple
columnCellTUnsafe row t = fromJust $ columnCellT row t

moveRow :: Int -> CellIndex -> CellIndex
moveRow val cell = let (RI row) = ciRowIndex cell
               in  cell { ciRowIndex = RI (row + val) }

moveColumn :: Int -> CellIndex -> CellIndex
moveColumn val cell = let (CI row) = ciColumnIndex cell
               in  cell { ciColumnIndex = CI (row + val) }

(+++) :: DT.Text -> RowIndex -> CellIndexTuple
(+++) = flip columnCellTUnsafe

(+^^) :: CellIndexTuple -> Int -> CellIndexTuple
(+^^) t i = tuple $ moveRow i $ fromTuple t

(+>>) :: CellIndexTuple -> Int -> CellIndexTuple
(+>>) t i = tuple $ moveColumn i $ fromTuple t

(^^^) :: CellIndexTuple -> RowIndex -> CellIndexTuple
(^^^) t row = let ci = fromTuple t
              in  tuple $ ci { ciRowIndex = row }

(>>>) :: CellIndexTuple -> DT.Text -> CellIndexTuple
(>>>) t col = let ci = fromTuple t
              in  tuple $ ci { ciColumnIndex = parseColumnIndexTextUnsafe col }

parseColumnIndexTextUnsafe :: DT.Text -> ColumnIndex
parseColumnIndexTextUnsafe = fromJust . parseColumnIndexText

parseCellIndexTextUnsafe :: DT.Text -> CellIndex
parseCellIndexTextUnsafe = fromJust . parseCellIndexText

{- |
  セルの値の取得.
-}
cellBoolValueFromSheet :: Worksheet -> CellIndex -> Bool
cellBoolValueFromSheet sheet cell = cellBoolValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellBoolValueFromSheetM :: Worksheet -> CellIndex -> Maybe Bool
cellBoolValueFromSheetM sheet cell = cellBoolValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellBoolValue :: Maybe CellValue -> Bool
cellBoolValue Nothing = error "cell not found"
cellBoolValue (Just (CellBool b)) = b
cellBoolValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellBoolValueM :: Maybe CellValue -> Maybe Bool
cellBoolValueM Nothing             = Nothing
cellBoolValueM (Just (CellBool b)) = Just b
cellBoolValueM (Just _)            = Nothing

cellDoubleValueFromSheet :: Worksheet -> CellIndex -> Double
cellDoubleValueFromSheet sheet cell = cellDoubleValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellDoubleValueFromSheetM :: Worksheet -> CellIndex -> Maybe Double
cellDoubleValueFromSheetM sheet cell = cellDoubleValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellDoubleValue :: Maybe CellValue -> Double
cellDoubleValue Nothing = 0
cellDoubleValue (Just (CellDouble d)) = d
cellDoubleValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellDoubleValueM :: Maybe CellValue -> Maybe Double
cellDoubleValueM Nothing               = Nothing
cellDoubleValueM (Just (CellDouble d)) = Just d
cellDoubleValueM (Just _)              = Nothing

cellIntegralValueFromSheet :: (Integral a) => Worksheet -> CellIndex -> a
cellIntegralValueFromSheet sheet cell = cellIntegralValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellIntegralValueFromSheetM :: (Integral a) => Worksheet -> CellIndex -> Maybe a
cellIntegralValueFromSheetM sheet cell = cellIntegralValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellIntegralValue :: (Integral a) => Maybe CellValue -> a
cellIntegralValue Nothing = 0
cellIntegralValue (Just (CellDouble d)) = ceiling d
cellIntegralValue (Just c)              = error ("unexpected cell value --> " ++ show c)

cellIntegralValueM :: (Integral a) => Maybe CellValue -> Maybe a
cellIntegralValueM Nothing               = Nothing
cellIntegralValueM (Just (CellDouble d)) = Just $ floor d
cellIntegralValueM (Just c)              = Nothing

cellStringValueFromSheet :: Worksheet -> CellIndex -> DT.Text
cellStringValueFromSheet sheet cell = cellStringValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellStringValueFromSheetM :: Worksheet -> CellIndex -> Maybe DT.Text
cellStringValueFromSheetM sheet cell = cellStringValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellStringValue :: Maybe CellValue -> DT.Text
cellStringValue Nothing = ""
cellStringValue (Just (CellText t)) = t
cellStringValue (Just c)            = error ("unexpected cell value --> " ++ show c)

cellStringValueM :: Maybe CellValue -> Maybe DT.Text
cellStringValueM Nothing             = Nothing
cellStringValueM (Just (CellText t)) = Just t
cellStringValueM (Just _)            = Nothing

cellDayValueFromSheet :: Worksheet -> CellIndex -> Day
cellDayValueFromSheet sheet cell = cellDayValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellDayValueFromSheetM :: Worksheet -> CellIndex -> Maybe Day
cellDayValueFromSheetM sheet cell = cellDayValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

cellDayValue :: Maybe CellValue -> Day
cellDayValue Nothing = error "cell not found"
cellDayValue mCell = let dayValue = cellIntegralValue mCell
                     in  addDays (dayValue - 2) $ fromGregorian 1900 1 1

cellDayValueM :: Maybe CellValue -> Maybe Day
cellDayValueM Nothing = Nothing
cellDayValueM mCell = let dayValue = cellIntegralValue mCell
                      in  Just $ addDays (dayValue - 2) $ fromGregorian 1900 1 1
