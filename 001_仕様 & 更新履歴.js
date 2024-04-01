/*
仕様

スプシをデータベースと同様に使用できるための機能を実装していく。


更新履歴
2024/02/21
スキーマの設定機能を実装。
文字の型を設定できるように。
デフォルトはString
設定するとInt、Jsonに対応。

2024/02/22
ユニークキーに設定されている値は保存前にチェックされ重複される場合はエラー処理。

2024/02/24
setValue()の仕様変更
Bool型 Date型 enumList型 set型に対応

2024/02/25
リレーション機能実装

2024/02/26
シート・キー名変換機能、実装。

2024/03/11
リッチテキストの保存を可能に。

2024/03/13
シートの保護機能を実装。
target座標の計算を追加

*/

/** 定数の説明 */

/**
 * 関数の説明
 * @param {int} arg1 - 引数の説明
 * @param {int} arg2 - 引数の説明
 * @return {bool} - 戻り値の説明
*/