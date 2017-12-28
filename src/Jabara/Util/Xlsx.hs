{-# LANGUAGE OverloadedStrings #-}
module Jabara.Util.Xlsx (
  RowIndex
  , ColumnIndex
  , SheetName
  , readBook
  , cellDoubleValue
  , cellDoubleValue'
  , cellIntegralValue
  , cellIntegralValue'
  , cellStringValue
  , cellStringValue'
) where

import           Codec.Xlsx
import           Control.Lens
import           Data.ByteString.Lazy (readFile)
import           Data.Maybe
import           Data.Text
import           Prelude              hiding (readFile)

readBook :: FilePath -> IO Xlsx
readBook path = readFile path >>= pure . toXlsx

type SheetName = Text
type RowIndex = Int
type ColumnIndex = Int

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

cellStringValue' :: (Integral a) => Worksheet -> (RowIndex, ColumnIndex) -> a
cellStringValue' sheet cell = cellIntegralValue $ sheet ^? ixCell cell . cellValue . _Just

cellStringValue :: Maybe CellValue -> Text
cellStringValue Nothing = ""
cellStringValue (Just (CellText t)) = t
cellStringValue (Just c)            = error ("unexpected cell value --> " ++ show c)

