import SheetDB_ from "./SheetDB";
// ライブラリ作成メモ
// https://tometome.notion.site/1630553833378060ae5eff818ad45118

var newDB = (spreadSheetId: string | null) => {
  return new SheetDB_(spreadSheetId);
};
