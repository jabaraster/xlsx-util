{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}
module Jabara.Xlsx.Types (
    RowIndex(..)

  , ColumnIndex(..)
  , parseColumnIndexText
  , formatColumnIndex

  , CellIndex(..)
  , parseCellIndexText
  , parseCellIndexTextUnsafe
  , cellUnsafe
  , cell
  , toTuple
  , moveRow
  , moveColumn
) where

import           Control.Exception
import           Data.Functor.Identity
import           Data.Maybe
import           Data.String           (IsString, fromString)
import qualified Data.Text             as DT (Text, foldr, length, pack, unpack)
import qualified Text.Parsec           as TP
import           Text.Parsec.Error     (ParseError (..))

{- \
  type RowIndex and instances.
-}
newtype RowIndex = RI {
    rowIndexValue :: Int
  }
  deriving (Eq, Ord)

instance Show RowIndex where
  show (RI i) = show (i + 1)

instance Read RowIndex where
  readsPrec i s = map (\(val::Int, remain) -> (RI val, remain)) $ readsPrec i s

instance Bounded RowIndex where
  minBound = RI 0
  maxBound = RI 1048575

instance Enum RowIndex where
  toEnum = RI
  fromEnum (RI i) = i

{- \
  type ColumnIndex and instances.
-}
newtype ColumnIndex = CI {
    columnIndexValue :: Int
  }
  deriving (Eq, Ord)

instance Show ColumnIndex where
  show = DT.unpack . formatColumnIndex

instance IsString ColumnIndex where
  fromString = fromJust . parseColumnIndexText . DT.pack

instance Bounded ColumnIndex where
  minBound = CI 0
  maxBound = CI 16383

instance Enum ColumnIndex where
  toEnum = CI
  fromEnum (CI i) = i

formatColumnIndex :: ColumnIndex -> DT.Text
formatColumnIndex (CI i)
  | i <= 25   = DT.pack [toEnum (i + 65)]
  | otherwise =
    DT.pack $ foldr foldCore ""
      $ addDigitLastElement
      $ collectDigitCount i []
  where
    foldCore :: Int ->  String -> String
    foldCore cnt text =
      let c::Char = toEnum (cnt + 65)
      in  c:text

addDigitLastElement is =
  let col = reverse $ map (flip (-) 1) $ tail $ reverse is
  in  col ++ [last is]

collectDigitCount :: Int -> [Int] -> [Int]
collectDigitCount 0 is = is
collectDigitCount remain is =
  let div = floor ((fromIntegral remain / 26)::Double)
      cnt = remain `mod` 26
  in  collectDigitCount div (cnt:is)

parseColumnIndexText :: DT.Text -> Maybe ColumnIndex
parseColumnIndexText t
  | DT.length t == 0 = Nothing
  | otherwise        = CI . flip (-) 1 . snd <$> DT.foldr core (Just (0, 0)) t
  where
    core :: Char -> Maybe (Int, Int) -> Maybe (Int, Int)
    core _ Nothing = Nothing
    core c (Just (digit, sum)) =
      let alpha = fromEnum c - 65 + 1
      in  if alpha < 1 || 26 < alpha
            then Nothing
            else
              let val = alpha * (26 ^ digit)
              in  Just (digit + 1, sum + val)

{- \
  type CellIndex and instances.
-}
data CellIndex =
  CellIndex {
    ciRowIndex    ::RowIndex
  , ciColumnIndex::ColumnIndex
  } deriving (Eq)

instance Show CellIndex where
  show ci = show (ciColumnIndex ci) ++ show (ciRowIndex ci)

instance Read CellIndex where
  readsPrec i s = case parseCellIndexText $ DT.pack $ snd $ splitAt i s of
    Nothing -> []
    Just ci -> [(ci, "")] -- 応用の効かない、よくない実装...

toTuple :: CellIndex -> (Int, Int)
toTuple ci = (rowIndexValue (ciRowIndex ci) + 1, columnIndexValue (ciColumnIndex ci) + 1)

cell :: DT.Text -> Maybe CellIndex
cell = parseCellIndexText

cellUnsafe :: DT.Text -> CellIndex
cellUnsafe = parseCellIndexTextUnsafe

moveRow ::Int -> CellIndex -> CellIndex
moveRow val cell = let (RI row) = ciRowIndex cell
               in  cell { ciRowIndex = RI (row + val) }

moveColumn ::Int -> CellIndex -> CellIndex
moveColumn val cell = let (CI row) = ciColumnIndex cell
               in  cell { ciColumnIndex = CI (row + val) }

parseCellIndexTextUnsafe :: DT.Text -> CellIndex
parseCellIndexTextUnsafe = fromJust . parseCellIndexText

parseCellIndexText :: DT.Text -> Maybe CellIndex
parseCellIndexText s = case parseCore s of
  Left _     -> Nothing
  Right cell -> Just cell
  where
    parseCore :: DT.Text -> Either ParseError CellIndex
    parseCore = TP.parse coreParser "" . DT.unpack

    coreParser :: TP.ParsecT String u Identity CellIndex
    coreParser = do
      colString <- TP.many1 TP.upper
      row <- rowParser
      mCi <- pure $ parseColumnIndexText $ DT.pack colString
      case mCi of
        Nothing -> error "parse fail."
        Just ci -> pure CellIndex {
                     ciRowIndex = row
                   , ciColumnIndex = ci
                   }

    rowParser :: TP.ParsecT String u Identity RowIndex
    rowParser = do
      rowString <- TP.many1 TP.digit
      pure $ RI $ flip (-) 1 $ read rowString
