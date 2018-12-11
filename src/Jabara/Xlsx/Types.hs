{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE TemplateHaskell     #-}
module Jabara.Xlsx.Types (
    RowIndex(..)
  , rowIndexValue

  , ColumnIndex(..)
  , columnIndexValue

  , CellIndex(..)
  , origin
  , ciRowIndex
  , ciColumnIndex
  , ciRow
  , ciColumn

  , CellIndexTuple

  , formatColumnIndex
  , parseColumnIndexText
  , parseColumnIndexTextUnsafe
  , parseCellIndexText
  , parseCellIndexTextUnsafe
) where

import           Control.Exception
import           Control.Lens
import           Control.Lens.TH
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
    _rowIndexValue :: Int
  }
  deriving (Eq, Ord)
makeLenses ''RowIndex

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
    _columnIndexValue :: Int
  }
  deriving (Eq, Ord)
makeLenses ''ColumnIndex

instance Show ColumnIndex where
  show = DT.unpack . formatColumnIndex

instance IsString ColumnIndex where
  fromString = parseColumnIndexTextUnsafe . DT.pack

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

    -- 26進数表現に於ける各桁毎の数値に変換する
    collectDigitCount :: Int -> [Int] -> [Int]
    collectDigitCount 0 is = is
    collectDigitCount remain is =
      let div = floor (fromIntegral remain / 26)
          cnt = remain `mod` 26
      in  collectDigitCount div (cnt:is)

parseColumnIndexText :: DT.Text -> Maybe ColumnIndex
parseColumnIndexText t
  | DT.length t == 0 = Nothing
  | otherwise        = CI . flip (-) 1 . snd <$> DT.foldr core (Just (0, 0)) t
  where
    core :: Char -> Maybe (Int, Int) -> Maybe (Int, Int)
    core _ Nothing = Nothing
    core c (Just (digit, sum))
      | fromEnum c `elem` [65..90] = Just (digit + 1, sum + (fromEnum c - 64) * (26 ^ digit))
      | otherwise = Nothing

parseColumnIndexTextUnsafe :: DT.Text -> ColumnIndex
parseColumnIndexTextUnsafe = fromJust . parseColumnIndexText

{- \
  type CellIndex and instances.
-}
data CellIndex =
  CellIndex {
    _ciRowIndex    :: RowIndex
  , _ciColumnIndex :: ColumnIndex
  } deriving (Eq)

{-| A1セル  -}
origin :: CellIndex
origin = CellIndex (RI 0) (CI 0)

makeLenses ''CellIndex

ciRow :: Lens' CellIndex RowIndex
ciRow f (CellIndex r c) = (`CellIndex` c) <$> f r

ciColumn :: Lens' CellIndex DT.Text
ciColumn f (CellIndex r c) = CellIndex r . parseColumnIndexTextUnsafe <$> f (formatColumnIndex c)

instance Bounded CellIndex where
  minBound = CellIndex minBound minBound
  maxBound = CellIndex maxBound maxBound

instance IsString CellIndex where
  fromString = parseCellIndexTextUnsafe . DT.pack

instance Show CellIndex where
  show = DT.unpack . formatCellIndex

instance Read CellIndex where
  -- 応用の効かない、よくない実装...
  readsPrec i s = maybe [] (\ci -> [(ci,"")]) $ parseCellIndexText $ DT.pack $ snd $ splitAt i s

formatCellIndex :: CellIndex -> DT.Text
formatCellIndex ci = DT.pack $ show (ci^.ciColumnIndex) ++ show (ci^.ciRowIndex)

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
                     _ciRowIndex = row
                   , _ciColumnIndex = ci
                   }

    rowParser :: TP.ParsecT String u Identity RowIndex
    rowParser = do
      rowString <- TP.many1 TP.digit
      pure $ RI $ flip (-) 1 $ read rowString

parseCellIndexTextUnsafe :: DT.Text -> CellIndex
parseCellIndexTextUnsafe = fromJust . parseCellIndexText
