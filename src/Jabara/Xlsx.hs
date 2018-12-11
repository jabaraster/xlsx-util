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

  , boolValue
  , boolValueM
  , boolCellValue
  , boolCellValueM

  , doubleValue
  , doubleValueM
  , doubleCellValue
  , doubleCellValueM

  , integralValue
  , integralValueM
  , integralCellValue
  , integralCellValueM

  , stringValue
  , stringValueM
  , stringCellValue
  , stringCellValueM

  , dayValue
  , dayValueM
  , dayCellValue
  , dayCellValueM

  , isEmpty
  , isEmpty'
  , traverseRow
  , traverseEveryRow

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
tuple :: CellIndex -> CellIndexTuple
tuple ci = (ci^.ciRowIndex^.rowIndexValue + 1, ci^.ciColumnIndex^.columnIndexValue + 1)

fromTuple :: CellIndexTuple -> CellIndex
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

columnCellT :: RowIndex -> DT.Text -> Maybe CellIndex
columnCellT row t = CellIndex row <$> parseColumnIndexText t

columnCellTUnsafe :: RowIndex -> DT.Text -> CellIndex
columnCellTUnsafe row t = fromJust $ columnCellT row t

moveRow :: Int -> CellIndex -> CellIndex
moveRow val cell =
  let (RI row) = cell^.ciRowIndex
  in  cell&ciRowIndex .~ RI (row + val)

moveColumn :: Int -> CellIndex -> CellIndex
moveColumn val cell =
  let (CI row) = cell^.ciColumnIndex
  in  cell&ciColumnIndex .~ CI (row + val)

(+++) :: ColumnIndex -> RowIndex -> CellIndex
(+++) = flip CellIndex 

{- |
  セルの値の取得.
-}
boolValue :: Worksheet -> CellIndex -> Bool
boolValue sheet cell = boolCellValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

boolValueM :: Worksheet -> CellIndex -> Maybe Bool
boolValueM sheet cell = boolCellValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

boolCellValue :: Maybe CellValue -> Bool
boolCellValue Nothing = error "cell not found"
boolCellValue (Just (CellBool b)) = b
boolCellValue (Just c)              = error ("unexpected cell value --> " ++ show c)

boolCellValueM :: Maybe CellValue -> Maybe Bool
boolCellValueM Nothing             = Nothing
boolCellValueM (Just (CellBool b)) = Just b
boolCellValueM (Just _)            = Nothing

doubleValue :: Worksheet -> CellIndex -> Double
doubleValue sheet cell = doubleCellValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

doubleValueM :: Worksheet -> CellIndex -> Maybe Double
doubleValueM sheet cell = doubleCellValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

doubleCellValue :: Maybe CellValue -> Double
doubleCellValue Nothing = 0
doubleCellValue (Just (CellDouble d)) = d
doubleCellValue (Just c)              = error ("unexpected cell value --> " ++ show c)

doubleCellValueM :: Maybe CellValue -> Maybe Double
doubleCellValueM Nothing               = Nothing
doubleCellValueM (Just (CellDouble d)) = Just d
doubleCellValueM (Just _)              = Nothing

integralValue :: (Integral a) => Worksheet -> CellIndex -> a
integralValue sheet cell = integralCellValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

integralValueM :: (Integral a) => Worksheet -> CellIndex -> Maybe a
integralValueM sheet cell = integralCellValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

integralCellValue :: (Integral a) => Maybe CellValue -> a
integralCellValue Nothing = 0
integralCellValue (Just (CellDouble d)) = ceiling d
integralCellValue (Just c)              = error ("unexpected cell value --> " ++ show c)

integralCellValueM :: (Integral a) => Maybe CellValue -> Maybe a
integralCellValueM Nothing               = Nothing
integralCellValueM (Just (CellDouble d)) = Just $ floor d
integralCellValueM (Just c)              = Nothing

stringValue :: Worksheet -> CellIndex -> DT.Text
stringValue sheet cell = stringCellValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

stringValueM :: Worksheet -> CellIndex -> Maybe DT.Text
stringValueM sheet cell = stringCellValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

stringCellValue :: Maybe CellValue -> DT.Text
stringCellValue Nothing = ""
stringCellValue (Just (CellText t)) = t
stringCellValue (Just c)            = error ("unexpected cell value --> " ++ show c)

stringCellValueM :: Maybe CellValue -> Maybe DT.Text
stringCellValueM Nothing             = Nothing
stringCellValueM (Just (CellText t)) = Just t
stringCellValueM (Just _)            = Nothing

dayValue :: Worksheet -> CellIndex -> Day
dayValue sheet cell = dayCellValue $ sheet ^? ixCell (tuple cell) . cellValue . _Just

dayValueM :: Worksheet -> CellIndex -> Maybe Day
dayValueM sheet cell = dayCellValueM $ sheet ^? ixCell (tuple cell) . cellValue . _Just

dayCellValue :: Maybe CellValue -> Day
dayCellValue Nothing = error "cell not found"
dayCellValue mCell = let dayValue = integralCellValue mCell
                     in  addDays (dayValue - 2) $ fromGregorian 1900 1 1

dayCellValueM :: Maybe CellValue -> Maybe Day
dayCellValueM Nothing = Nothing
dayCellValueM mCell = let dayValue = integralCellValue mCell
                      in  Just $ addDays (dayValue - 2) $ fromGregorian 1900 1 1


{-|
  ユーティリティ
-}

isEmpty :: Worksheet -> CellIndex -> Bool
isEmpty sheet cell = maybe True core $ sheet ^? ixCell (tuple cell) . cellValue . _Just
  where
    core :: CellValue -> Bool
    core (CellText t) = t == ""
    core _            = False

isEmpty' :: Worksheet -> ColumnIndex -> RowIndex -> Bool
isEmpty' sheet col row = isEmpty sheet (CellIndex row col)

traverseRow :: Worksheet -> RowIndex -> Int -> (RowIndex -> Bool) -> (RowIndex -> Maybe a) -> [a]
traverseRow sheet startRow skip endDecider converter =
  core sheet startRow skip endDecider converter []
  where
    core:: Worksheet -> RowIndex -> Int -> (RowIndex -> Bool) -> (RowIndex -> Maybe a) -> [a] -> [a]
    core sheet currentRow skip endDecider converter collected =
      if endDecider currentRow
        then collected
        else
          let newCollected = maybe collected (\val -> collected ++ [val]) $ converter currentRow
          in core sheet (currentRow +.. skip) skip endDecider converter newCollected

traverseEveryRow :: Worksheet -> RowIndex -> (RowIndex -> Bool) -> (RowIndex -> Maybe a) -> [a]
traverseEveryRow sheet startRow = traverseRow sheet startRow 1
