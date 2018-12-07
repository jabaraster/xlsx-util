{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}
module Jabara.Xlsx.Types (
    RowIndex(..)
  , ColumnIndex(..)
  , CellIndex(..)

  , CellIndexTuple

  , formatColumnIndex
  , parseColumnIndexText
  , parseCellIndexText
) where

import           Control.Exception
import           Data.Functor.Identity
import           Data.Maybe
import           Data.String           (IsString, fromString)
import qualified Data.Text             as DT (Text, foldr, length, pack, unpack)
import qualified Text.Parsec           as TP
import           Text.Parsec.Error     (ParseError (..))

type CellIndexTuple = (Int, Int)

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

{- \
  type CellIndex and instances.
-}
data CellIndex =
  CellIndex {
    ciRowIndex    ::RowIndex
  , ciColumnIndex::ColumnIndex
  } deriving (Eq)

instance Show CellIndex where
  show = DT.unpack . formatCellIndex

instance Read CellIndex where
  -- 応用の効かない、よくない実装...
  readsPrec i s = maybe [] (\ci -> [(ci,"")]) $ parseCellIndexText $ DT.pack $ snd $ splitAt i s

formatColumnIndex :: ColumnIndex -> DT.Text
formatColumnIndex (CI i)
  | i <= 25   = DT.pack [toEnum (i + 65)]
  | otherwise =
    DT.pack $ foldr foldCore ""
      $ adjustDigit
      $ collectDigitCount i []
  where
    foldCore :: Int -> String -> String
    foldCore cnt text =
      let c::Char = toEnum (cnt + 65)
      in  c:text

    -- Excelのカラムのテキストは1桁目だけルールがおかしい.
    -- 1桁目は1=A,26=Zだが、2桁目以上は0=A,25=Zとなっている.
    -- これを補正する.
    adjustDigit is =
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

formatCellIndex :: CellIndex -> DT.Text
formatCellIndex ci = DT.pack $ show (ciColumnIndex ci) ++ show (ciRowIndex ci)

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
