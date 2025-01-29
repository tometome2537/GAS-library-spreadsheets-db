class SheetDB_ {
  // コンストラクタ
  constructor(spreadSheetId) {
    // スプシID初期値
    this._spreadSheetId = null;
    // スプシID初期値の定義を実行
    this.spreadSheetId = spreadSheetId ? spreadSheetId : null;

    // スキーマを定義
    this._schema = null;
    // シート名・キー名変更履歴を定義
    this._historySheetKeyName = null;

    // スプシのメタ情報を保存するシート名
    // python等の別のライブラリで読み込むこともあるため値の変更は非推奨。
    this._metaDataSheetName = "_metadata";

    // 空白のセルをnullとして扱うかどうか
    // To Do

    // プロテクションしたシート
    this._protectionSheetNames = [];

    // キャッシュのリセットを実行
    this.cacheReset();
  }

  /** キャッシュのリセット */
  cacheReset() {
    // スプシファイルの初期化
    this._spreadSheet = null;
    // 全シートのシートのキャッシュを初期化
    this._cacheSheetAllSheet = {};
    // 全シートの値①のキャッシュを初期化
    this._cacheSheetValuesAllSheet = {};
    // 全シートの値②のキャッシュを初期化
    // 全シートの値③のキャッシュを初期化
    // 全シートの値④のキャッシュを初期化

    // sheetObjのキャッシュ
    this._cacheSheetObj = {};

    // 保存するy座標のキャッシュ 繰り返し処理でこの関数を使用すると同じy座標の位置に値が保存されるのを防止するため。
    this._ySetValueAppEndRow = null;

    // ↓ ユニークキーのチェックをしたシートの名前が入る配列
    this._cacheUniqueKeyDone = [];
    // ↓ ユニークキーの値の配列
    this._cacheUniqueKeyValues = {};
  }

  /**
   * スプレッドシートIDをセットする
   * @param {string} spreadSheetId - スプシIDを定義
   */
  set spreadSheetId(spreadSheetId) {
    if (spreadSheetId) {
      this._spreadSheetId = spreadSheetId;
    } else {
      this._spreadSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    }
  }

  /** スプレッドシートIDを取得 */
  get spreadSheetId() {
    if (this._spreadSheetId) {
      return this._spreadSheetId;
    } else {
      throw '読み込むスプシIDが定義されていません。';
    }
  }

  /** スプシファイルを取得 */
  get spreadSheet() {
    if (this._spreadSheet) {
      // キャッシュが保存されている場合。
      return this._spreadSheet;
    } else {
      const spreadSheet = SpreadsheetApp.openById(this.spreadSheetId);
      this._spreadSheet = spreadSheet;
      return spreadSheet;
    }
  }

  /** シート名を一覧を取得 */
  getSheetNames() {
    return this.spreadSheet.getSheets().map(sheet => sheet.getName());
  }
  /** メタデータ保存のシート名を取得 */
  get metaDataSheetName() {
    // メタデータシートが存在しない場合は
    if (!(this.getSheetNames().includes(this._metaDataSheetName))) {
      // シートの作成。
      let newSheet = this.spreadSheet.insertSheet(this._metaDataSheetName);
      newSheet.appendRow(["key", "value"]);
      // シートをアクティベート。
      this.spreadSheet.setActiveSheet(newSheet);
    }
    return this._metaDataSheetName;
  }
  // スキーマの保存
  set schema(schema) {
    // スキーマをmetaDataシートに保存する。(ただし既存の値を変更点がない場合は何もしない。)
    const metaDataSheet = this.getSheetByName(this.metaDataSheetName);
    const cellData = metaDataSheet.getRange("B2").getValue();
    if (JSON.stringify(schema) !== cellData) {
      metaDataSheet.getRange('A2').setValue("schema");
      metaDataSheet.getRange("B2").setValue(JSON.stringify(schema));
    }
    // classプロパティに保存
    this._schema = schema;
  }
  // スキーマの読み取り
  get schema() {
    // スキーマが事前に渡されている場合
    if (this._schema) {
      return this._schema;
    }
    // スキーマがclassプロパティにない場合
    const metaDataSheet = this.getSheetByName(this.metaDataSheetName);
    const cellData = metaDataSheet.getRange("B2").getValue();
    const schema = cellData === "" ? {} : JSON.parse(cellData);
    this._schema = schema;
    return schema;
  }
  // シート名・キー名変更履歴をセット。
  set historySheetKeyName(historySheetKeyName) {
    if (this._historySheetKeyName) {
      throw `シート・キー履歴が再度設定されようとしています。`;
    } else {
      // 値をmetaDataシートに保存する。(ただし既存の値を変更点がない場合は何もしない。)
      const metaDataSheet = this.spreadSheet.getSheetByName(this.metaDataSheetName);
      const cellData = metaDataSheet.getRange("B3").getValue();
      if (JSON.stringify(historySheetKeyName) !== cellData) {
        metaDataSheet.getRange('A3').setValue("historySheetKeyName");
        metaDataSheet.getRange("B3").setValue(JSON.stringify(historySheetKeyName));
      }
      // classプロパティに保存
      this._historySheetKeyName = historySheetKeyName;
    }
  }
  // シート名・キー名変更履歴を呼び出し。
  get historySheetKeyName() {
    if (this._historySheetKeyName) {
      return this._historySheetKeyName;
    }
    // シート名・キー名変更履歴がclassプロパティにない場合
    const metaDataSheet = this.spreadSheet.getSheetByName(this.metaDataSheetName);
    const cellData = metaDataSheet.getRange("B3").getValue();
    const historySheetKeyName = cellData === "" ? {} : JSON.parse(cellData);
    this._historySheetKeyName = historySheetKeyName;
    return historySheetKeyName;
  }
  // 最新のシート名の履歴を取得する。
  getLatestSheetName(sheetName) {
    if (!sheetName) {
      throw 'シート名が定義されていません。';
    }
    for (const historySheetName in this.historySheetKeyName) {
      if ('sheetNameHistory' in this.historySheetKeyName[historySheetName]) {
        if (
          this.historySheetKeyName[historySheetName][
            'sheetNameHistory'
          ].includes(sheetName)
        ) {
          return historySheetName;
        }
      }
    }
    return sheetName;
  }
  /** キー名を取得 */
  getLatestKeyName(sheetName, keyName) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    //キャッシュに登録されているシート名かチェック
    if ([latestSheetName] in this.historySheetKeyName) {
      // 保存されている場合
      if ('keyNameHistory' in this.historySheetKeyName[latestSheetName]) {
        // keyNameHistoryの項目があるかどうか
        if (
          [keyName] in
          this.historySheetKeyName[latestSheetName]['keyNameHistory']
        ) {
          return this.historySheetKeyName[latestSheetName]['keyNameHistory'][
            keyName
          ];
        }
      }
    }
    // 保存されていない場合
    return keyName;
  }
  /**
   * シートを読み込む
   * @param {string} sheetName - シートの名前
   * @return {function} - シート
   */
  getSheetByName(sheetName) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    if ([latestSheetName] in this._cacheSheetAllSheet) {
      // キャッシュに保存されている場合
      return this._cacheSheetAllSheet[latestSheetName];
    } else {
      // キャッシュに保存されていない場合
      const sheet = this.spreadSheet.getSheetByName(latestSheetName);
      // キャッシュに保存
      this._cacheSheetAllSheet[latestSheetName] = sheet;
      return sheet;
    }
  }

  // シートの値①(関数の結果)を読み込む
  getSheetValues(sheetName) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    let result = null;
    if ([latestSheetName] in this._cacheSheetValuesAllSheet) {
      // キャッシュを返す
      result = this._cacheSheetValuesAllSheet[latestSheetName];
    } else {
      const sheetValues = this.getSheetByName(latestSheetName)
        .getDataRange()
        .getValues(); // 値を取得(推測で型変換あり)
      this._cacheSheetValuesAllSheet[latestSheetName] = sheetValues; // キャッシュに保存
      result = sheetValues;
    }
    return SheetDB_.deepCopy(result);
  }
  // シートの値②(元の関数)を読み込む ※ 関数のみ取得。関数でないセルは ""(空白)で取得される。 (例 =sum() )
  getSheetFormulas(sheetName) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    return this.getSheetByName(latestSheetName).getDataRange().getFormulas();
  }
  // シートの値③(ディスプレイの表示される値) 0.01ではなく1%で取得される。すべて文字型で出力される。
  getSheetDisplayValues(sheetName) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    return this.getSheetByName(latestSheetName)
      .getDataRange()
      .getDisplayValues();
  }
  // シートの値④(リッチテキスト)(フォント、文字色、太さ etc...)
  getSheetRichTextValues(sheetName) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    return this.getSheetByName(latestSheetName)
      .getDataRange()
      .getRichTextValues();
  }
  /**
   * シートのobjを取得
   * @param {string} sheetName - シートの名前
   * @param {int} relationCount - リレーションの深さ(0でリレーションなし。)
   * @param {Object} targets - targetを指定する
   * @param {string} dataType - 値の種類(schema、string、values①、②、③、④)
   * 空白のセルをnullとして扱うかどうか。
   * @return {Obj} - Obj(Json)
   */
  getSheetObj(sheetName, relationCount, targets, dataType) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);
    // レスポンスする関数
    let sheetObj = null; // obj初期化

    // キャッシュ対象かどうかを判定(Bool値)
    const cacheEligible = !relationCount && !targets && !dataType;

    // キャッシュ対象の場合、キャッシュに値があればそこから返す。
    if (cacheEligible) {
      if (this._cacheSheetObj[sheetName]) {
        return this._cacheSheetObj[sheetName];
      }
    }

    // データタイプ
    if (dataType === 'schema' || !dataType) {
      // 値を取得
      const sheetValues = this.getSheetValues(latestSheetName);

      // objに変換
      sheetObj = SheetDB_.convertArrayToObjectString(sheetValues);

      // スキーマを定義
      const schema = this.schema[latestSheetName];

      // ⭐️ここからスキーマの通りに変換
      if (schema) {
        // スキーマが定義されている場合
        for (let item of sheetObj) {
          for (const key in schema) {
            // 定義を繰り返す
            if ('dataType' in schema[key]) {
              // BigInt型の場合 桁数の多い数字は文字型として出力する。デフォルトで桁数が多い数値は文字列としてスプシ(GAS)は扱っている。
              if (schema[key]['dataType'].match(/bigint/i)) {
                if (typeof item[key] === 'number') {
                  // 桁数の少ない数値は数値型なので統一させるために文字型として出力する。
                  item[key] = String(item[key]);
                } else if (typeof item[key] === 'string') {
                  // 何もしない(デフォルトの文字型の場合はそのまま出力)
                } else {
                  item[key] = null;
                }

                // 数値型の場合
              } else if (schema[key]['dataType'].match(/int/i)) {
                // 空文字をNumber()すると0になってしまう。 undefinedを返すとJSON.stringify()でkeyが削除されてしまうためnullを返す。
                if (item[key] === '' || item[key] === null) {
                  // nullはNaNに変換
                  item[key] = NaN;
                } else {
                  item[key] = Number(item[key]);
                }
              }
              // jsonの場合
              if (schema[key]['dataType'].match(/json/i)) {
                try {
                  item[key] = JSON.parse(item[key]);
                } catch {
                  item[key] = null;
                }
              }
              // Bool型の場合
              if (schema[key]['dataType'].match(/bool|boolen|boolean/i)) {
                if (item[key] === null || item[key] === '') {
                  item[key] = null;
                } else if (/^(true|yes|1)$/i.test(item[key])) {
                  item[key] = true;
                } else if (/^(false|no|0)$/i.test(item[key])) {
                  item[key] = false;
                } else {
                  item[key] = null;
                }
              }
              // date型の場合
              if (schema[key]['dataType'].match(/date/i)) {
                item[key] = new Date(item[key]);
                if (item[key].toString() === 'Invalid Date') {
                  // 日付に変換できなかった場合はnull
                  item[key] = null;
                }
              }
              // enumList(重複が許される)の場合
              if (schema[key]['dataType'].match(/enumlist|set/i)) {
                if (typeof item[key] === 'string' && item[key].length >= 1) {
                  item[key] = item[key].split(/ , |,| ,|, /).filter(v => v);
                  // Set型(重複が許されない)
                  if (schema[key]['dataType'].match(/set/i)) {
                    item[key] = item[key].filter(
                      (item, index, self) => self.indexOf(item) === index
                    );
                  }
                } else {
                  item[key] = null;
                }
              }

              // String型の扱い
              if (schema[key]['dataType'].match(/string/i)) {
                if (typeof item[key] === 'string' && item[key].length >= 1) {
                  // １文字以上の文字列なら文字型に変換
                  item[key] = String(item[key]);
                } else {
                  // それ以外はnull
                  item[key] = null;
                }
              } else {
                item[key] = item[key];
              }
            }
          }
        }
      }
      // デフォルトはnullにする。
      for (let sheetObjItem of sheetObj) {
        for (let sheetObjItemKey in sheetObjItem) {
          // デフォルトはnull
          if (sheetObjItem[sheetObjItemKey] === "") {
            sheetObjItem[sheetObjItemKey] = null;
          }
        }
      }
      // ⭐️ここまでスキーマの通りに変換
    } else if (dataType === 'string') {
      // 値を取得
      const sheetValues = this.getSheetValues(latestSheetName);
      // objに変換
      sheetObj = SheetDB_.convertArrayToObjectString(sheetValues);
    } else if (dataType === 'values') {
      // 値を取得
      const sheetValues = this.getSheetValues(latestSheetName);
      // objに変換
      sheetObj = SheetDB_.convertArrayToObject(sheetValues);
    }

    // ターゲットが設定されている。
    if (targets) {
      // 座標を取得
      const tergetCoordinate = this.getTargetCoordinate(
        latestSheetName,
        targets
      );

      // yの値のみを配列にする。
      const yCoordinate = tergetCoordinate.map(v => v['y']);

      sheetObj = sheetObj.filter((resultItem, index) => {
        return yCoordinate.includes(index + 2);
      });
    }

    // リレーション
    if (relationCount === 0 || !relationCount) {
      // 何もしない。
    } else {
      // スキーマを定義
      const schema = this.schema[latestSheetName];
      if (schema) {
        // スキーマでシートが定義されていれば。
        for (const key of Object.keys(schema)) {
          // シートの各keyの定義を繰り返す
          if ('relation' in schema[key]) {
            // リレーションの存在を確認
            // リレーション先のobjを取得
            const relationSheetObj = this.getSheetObj(
              schema[key]['relation']['sheetName'],
              relationCount - 1,
              null, // リレーション2層目以降はtarget引数は適応しない。
              dataType
            );
            // リザルトに付与
            for (let resultItem of sheetObj) {
              resultItem[key] = relationSheetObj.filter(relationObj => {
                // リレーション元のkeyのvalueが配列[SET型][enumList型]
                if (
                  Array.isArray(
                    resultItem[schema[key]['relation']['references']]
                  )
                ) {
                  if (
                    relationObj[schema[key]['relation']['sheetKey']] &&
                    resultItem[schema[key]['relation']['references']]
                  ) {
                    // null以外の場合

                    return resultItem[
                      schema[key]['relation']['references']
                    ].includes(
                      relationObj[schema[key]['relation']['sheetKey']]
                    );
                  }
                }

                // リレーション先のkeyのvalueが配列[SET型][enumList型]
                if (
                  Array.isArray(
                    relationObj[schema[key]['relation']['sheetKey']]
                  )
                ) {
                  if (
                    relationObj[schema[key]['relation']['sheetKey']] &&
                    resultItem[schema[key]['relation']['references']]
                  ) {
                    // null以外の場合

                    return relationObj[
                      schema[key]['relation']['sheetKey']
                    ].includes(
                      resultItem[schema[key]['relation']['references']]
                    );
                  }
                }

                // リレーション先のkeyのvalueが文字列(object[SET型]以外)の場合は型を含めた完全一致で定義する。

                return (
                  relationObj[schema[key]['relation']['sheetKey']] ===
                  resultItem[schema[key]['relation']['references']]
                );
              });
            }
          }
        }
      }
    }



    // レスポンス前にキャッシュ対象の場合はキャッシュに値を保存する。
    if (cacheEligible) {
      this._cacheSheetObj[sheetName] = sheetObj;
    }

    // レスポンス
    return SheetDB_.deepCopy(sheetObj);
  }

  // X軸を調べる
  getXCoordinate(sheetName, xKey) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    // xKeyは必ず文字型で受け取ること。0の場合は"0"で受け取る。
    if (!(typeof xKey === 'string' && xKey.length >= 1)) {
      throw `xKeyは文字型で１文字以上の必要があります。${latestSheetName}の${xKey}`;
    }

    // ↓ １行目の値を配列を文字型で取得。
    const xKeyList = this.getSheetValues(latestSheetName)[0].map(String);

    let result = []; // 結果を初期化
    let count = 1; // 初期値

    // ↓ 取得した値からxKeyは何番目にあるかを計算
    for (const xKeyListItem of xKeyList) {
      if (xKeyListItem === xKey) {
        result.push(count);
      }
      count += 1; // ⇦xKeyは何番目にあるかをカウントするための変数。もっといい書き方ある説。
    }
    if (result.length >= 2) {
      console.warn(
        `${latestSheetName}に${xKey}キーが２つ以上あります。プログラムが意図した通りに動作しない可能性があります。`
      );
      // throw `${latestSheetName}に${xKey}キーが２つ以上あります。`;
    }
    return result;
  }

  // Y座標を取得する
  getYCoordinate(sheetName, xKey, value) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    // ↓ false、0、NaN、null、""でy座標を探す場合も考えられるのでこの書き方。
    if (value === undefined) {
      throw `valueが空白です!!`;
    }
    // シートを取得
    const sheetObj = this.getSheetObj(latestSheetName);

    let result = {}; // 結果の初期化

    for (const x of this.getXCoordinate(latestSheetName, xKey)) {
      // xキー番号を繰り返す
      let count = 2; // 初期値
      result[x] = []; // 初期化

      for (const sheetObjItem of sheetObj) {
        // 両方とも配列だった場合(AND検索)
        if (
          sheetObjItem[xKey] &&
          Array.isArray(sheetObjItem[xKey]) &&
          Array.isArray(value)
        ) {
          if (
            value.every(valueItem => sheetObjItem[xKey].includes(valueItem))
          ) {
            result[x].push(count);
          }

          // 片方が配列で片方が文字型だった場合
        } else if (
          sheetObjItem[xKey] &&
          Array.isArray(sheetObjItem[xKey]) &&
          typeof value === 'string'
        ) {
          if (sheetObjItem[xKey].includes(value)) {
            result[x].push(count);
          }

          // 片方が配列で片方が文字型だった場合
        } else if (
          sheetObjItem[xKey] &&
          typeof sheetObjItem[xKey] === 'string' &&
          Array.isArray(value)
        ) {
          if (value.includes(sheetObjItem[xKey])) {
            result[x].push(count);
          }

          //両方をともNaNだった場合
        } else if (
          typeof sheetObjItem[xKey] === 'number' &&
          isNaN(sheetObjItem[xKey]) &&
          typeof value === 'number' &&
          isNaN(value)
        ) {
          result[x].push(count);

          // その他(文字型・Boolean型・null・"")
        } else if (sheetObjItem[xKey] === value) {
          result[x].push(count);

          // BigInt型に対応(Valueが桁数の多い数値の場合に文字型に変換String()すると値が変わってしまうためシートデータ側をNumber()して比較する。))
        } else if (Number(sheetObjItem[xKey]) === value) {
          result[x].push(count);
        }
        count += 1;
      }
    }

    return result;
  }

  // target座標を取得する
  getTargetCoordinate(sheetName, targets) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    /* 
    ・targetsの引数に以下のような同一のkey(この場合person)を持った配列が入力された場合エラー処理する。 → JavaScript言語の仕様で多分無理。
    {"person": "名前１", "person": "名前２"}
    以下のようにすれば意図した動作が可能
    { "person": ["名前１", "名前２"] }
    */

    // 結果を初期化
    let result = [];

    // xの値を削除する必要性があるかどうか
    let shouldRemoveX = false;

    // targetを繰り返す → targetが2つ以上指定されている → targetの条件をすべて満たした座標(AND検索)をreturnする必要がある。
    for (const target in targets) {
      const targetXKey = target;
      const targetYKey = targets[target];

      // 座標を調べる
      const xs = this.getXCoordinate(latestSheetName, targetXKey);
      const ys = this.getYCoordinate(latestSheetName, targetXKey, targetYKey);

      // 座標をresultに蓄積
      if (result.length === 0) {
        // targets１周目
        for (const x of xs) {
          // X軸を繰り返す
          for (const y of ys[x]) {
            // y軸を繰り返す
            const obj = { x: x, y: y };
            result.push(obj);
          }
        }
      } else {
        // targets２週目以降
        // 結果をキャッシュに保存
        const cache = result;
        // 初期化して
        result = [];

        for (const x of xs) {
          // X軸を繰り返す
          // xの値を削除する必要性があるかどうかの確認
          if (!shouldRemoveX) {
            if (cache[0]['x'] !== x) {
              shouldRemoveX = true;
            }
          }
          for (const y of ys[x]) {
            // y軸を繰り返す
            for (const cacheItem of cache) {
              // キャッシュを繰り返す
              if (cacheItem['y'] === y) {
                const obj = { x: x, y: y };
                result.push(obj);
              }
            }
          }
        }
      }
    }

    // xの値を削除する必要性があれば
    if (shouldRemoveX) {
      // xの値を削除
      result = result.map(({ x, ...rest }) => rest);
    }

    // 直近のターゲットとして保存する。
    this._targetCoordinate = result;

    // 結果を返す
    return result;
  }

  // リッチテキストの保存
  setRichTextValue(sheetName, targets, newRichTextValue) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);
    // ターゲティング
    const targetCoordinate = this.getTargetCoordinate(latestSheetName, targets);
    // targetを繰り返す
    for (const target of targetCoordinate) {
      // 自分自身の値の置き換えが可能になるためのセーフティ機能必要
      if (
        newRichTextValue.getText() !=
        this.getCellValue(latestSheetName, target['y'], target['x'])['values']
      ) {
        throw '既存の値とsetText()で定義された値が違います。';
      } else {
        // 設定
        this.getSheetByName(latestSheetName)
          .getRange(target['y'], target['x'])
          .setRichTextValue(newRichTextValue);
      }
    }
  }

  // 値を保存
  setValue(sheetName, targets, setData) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);
    // targetを定義
    const targetCoordinate = this.getTargetCoordinate(latestSheetName, targets);
    if (targetCoordinate.length === 0) {
      console.error('targetが0です。値の保存は実行されません。');
    }
    let result = []; // 初期化
    // 座標のy座標を繰り返す。
    for (const target of targetCoordinate) {
      // 保存を実行
      result.push(this.setValueDone_(latestSheetName, target['y'], setData));
    }
    return result;
  }
  // 値の保存(行の一番下)
  setValueAppEndRow(sheetName, setData) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    // 値を保存するy座標が定義されていない場合は定義する
    if (!this._ySetValueAppEndRow) {
      // 値を保存するy座標を検出
      this._ySetValueAppEndRow =
        this.getSheetValues(latestSheetName).length + 1;
    }


    // レスポンスを初期化
    const response = [];

    // 保存を実行
    const result = this.setValueDone_(latestSheetName, this._ySetValueAppEndRow, setData)

    response.push(
      result
    );

    // 値を保存するy座標の値を更新
    if (result["status"] === "success") {
      // 保存に成功していた場合
      this._ySetValueAppEndRow += 1;
    } else if (result["status"] === "error") {
      // 保存に失敗していた場合
      // 何もしない
    } else {
      // それ以外の場合
      throw "setValueAppEndRow()の実行において保存の成功可否が不明です。"
    }


    return response;
  }
  // 保存を実行(内部関数)
  setValueDone_(sheetName, y, setData) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    // 保存を実行するかどうか(初期化)
    let setValueDone = true;
    // エラーメッセージ(初期化)
    let errorText = '';
    // ユニークキーを保存
    this.cacheUniqueKey_(latestSheetName);

    // シートを取得
    const sheet = this.getSheetByName(latestSheetName);

    // X座標が存在するかチェック
    for (const key in setData) {
      // X座標取得しようとしてみる。(存在しないkeyの場合はエラー)
      if (this.getXCoordinate(latestSheetName, key).length === 0) {
        setValueDone = false;
        errorText += 'setDataに存在しないidが入力されました。';
      }
    }
    // 保存していいかユニークキーチェック
    for (const key in setData) {
      if (latestSheetName in this._cacheUniqueKeyValues) {
        if ([key] in this._cacheUniqueKeyValues[latestSheetName]) {
          // キャッシュにキーが存在するか確認
          if (this._cacheUniqueKeyValues[latestSheetName][key].length !== 0) {
            if (
              this._cacheUniqueKeyValues[latestSheetName][key].includes(
                setData[key]
              )
            ) {
              setValueDone = false;
              errorText += 'ユニークキーとしてすでに値が保存されています。';
            }
          }
        }
      }
    }

    // シートをロックする
    if (!this._protectionSheetNames.includes(latestSheetName)) {
      this.protectionSheet(latestSheetName);
    }

    // 保存を実行していい場合。
    if (setValueDone) {
      // 初期化
      let result = {
        "status": "success",
        "value": {}
      };
      // setDataを繰り返して保存を実行
      for (const key in setData) {
        // X座標を取得
        const xs = this.getXCoordinate(latestSheetName, key);

        // x座標を繰り返す
        setValueLoop: for (const x of xs) {
          // すでに保存されている値を定義
          const oldValue = this.getCellValue(latestSheetName, y, x)['values'];
          let newValue = null;

          // 保存を実行
          // SET型またはenumlist型に指定されている場合。
          if (
            Array.isArray(setData[key]) &&
            [key] in this.schema[latestSheetName] &&
            this.schema[latestSheetName][key]['dataType'].match(/enumlist|set/i)
          ) {
            // 重複を許可しない。(set型)
            if (/set/i.test(this.schema[latestSheetName][key]['dataType'])) {
              newValue = [...new Set(setData[key])].join(' , ');
              // 重複を許可(enumlist型)
            } else {
              newValue = setData[key].join(' , ');
            }

            if (oldValue === newValue) {
              continue setValueLoop;
            } else {
              sheet.getRange(y, x).setValue(newValue);
            }

            // Bool型の場合
          } else if (typeof setData[key] === 'boolean') {

            // 新しい値を文字列
            newValue = setData[key];


            if (oldValue === newValue) {
              continue setValueLoop;
            } else {
              sheet.getRange(y, x).setValue(newValue);
            }


            //json(obj)の場合
          } else if (typeof setData[key] === 'object') {
            newValue = JSON.stringify(setData[key]);
            if (oldValue === newValue) {
              continue setValueLoop;
            } else {
              sheet.getRange(y, x).setValue(newValue);
            }
          } else {
            // その他の場合 数値型と文字型を比較することもあるため == である。
            newValue = setData[key];
            if (oldValue == newValue) {
              continue setValueLoop;
            } else {
              sheet.getRange(y, x).setValue(newValue);
            }
          }

          // ユニークキーをキャッシュに追記
          if (latestSheetName in this._cacheUniqueKeyValues) {
            if ([key] in this._cacheUniqueKeyValues[latestSheetName]) {
              //キャッシュにキーが存在する場合
              this._cacheUniqueKeyValues[latestSheetName][key].push(
                setData[key]
              );
            }
          }

          // retrun result用objの追加
          result["value"][key] = { oldValue: oldValue, newValue: newValue };
        } // xsのforの閉じタグ
      } // setDataのfor閉じタグ
      return result;
    } else {
      return {
        "status": "error",
        "errorMessage": errorText
      }
    }
  }

  // ユニークキーを保存する
  cacheUniqueKey_(sheetName) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);
    // スキーマを定義
    const schema = this.schema[latestSheetName];

    // ユニークキーが保存処理が未実行の場合　かつ スキーマが定義されている場合。
    if (!this._cacheUniqueKeyDone.includes(latestSheetName) && schema) {
      for (const key of Object.keys(schema)) {
        // シートの各keyの定義を繰り返す
        if ('decorator' in schema[key]) {
          // decoratorの存在を確認
          if (schema[key]['decorator'].includes('unique')) {
            // decoratorにuniqueが指定されている場合
            // 既存の値をキャッシュにすべて保存する
            this._cacheUniqueKeyValues[latestSheetName] = {}; // 初期化
            this._cacheUniqueKeyValues[latestSheetName][key] = []; // 初期化
            for (const item of this.getSheetObj(latestSheetName)) {
              this._cacheUniqueKeyValues[latestSheetName][key].push(item[key]);
            }
          }
        }
      }
      this._cacheUniqueKeyDone.push(latestSheetName);
    }
  }

  // 特定のセルの値を調べる
  getCellValue(sheetName, y, x) {
    // シート名を履歴から呼び出す。
    const latestSheetName = this.getLatestSheetName(sheetName);

    let result = {};
    if (this.getSheetValues(latestSheetName).length < y) {
      result['values'] = undefined;
    } else {
      result['values'] = this.getSheetValues(latestSheetName)[y - 1][x - 1];
    }

    return result;
  }

  // シートを保護するコード
  protectionSheet(sheetNames) {
    if (typeof sheetNames === 'string') {
      sheetNames = [sheetNames];
    }

    for (let sheetName of sheetNames) {
      // シート名を履歴から呼び出す。
      const latestSheetName = this.getLatestSheetName(sheetName);

      const protection = this.getSheetByName(latestSheetName)
        .protect()
        .setDescription('GASの更新処理中のため保護');
      // 編集時に警告を表示する
      protection.setWarningOnly(true);
      // 保存
      this._protectionSheetNames.push(latestSheetName);
    }
  }

  // シートの保護を解除するコード
  protectionRemoveSheet() {
    // シート名を履歴から呼び出す。
    // const latestSheetName = this.getLatestSheetName(sheetName)

    for (const sheetName of this.getSheetNames()) {
      const protection = this.getSheetByName(sheetName).getProtections(
        SpreadsheetApp.ProtectionType.SHEET
      )[0];
      if (protection) {
        protection.remove();
      }
    }
  }

  // シートを入れるとobj形式に変換してくれる。
  static convertArrayToObject(sheetValues) {
    const rows = SheetDB_.deepCopy(sheetValues); // deepCopy
    const keys = rows.splice(0, 1)[0];
    return rows.map(row => {
      const obj = {};
      row.map((item, index) => {
        // 保存されている値が0の場合にfalse判定になりnullが出力されてしまうのでString(item) === ""の記述で判定を行っている。
        // obj[String(keys[index])] = String(item) === "" ? null : String(item);
        // すべての値を文字型にして出力する方が関数としての役割。JavaScript上では扱いやすい。
        obj[String(keys[index])] = item;
      });
      return obj;
    });
  }
  // シートを入れるとobj形式に変換してくれる。値はすべて文字型に変換。
  static convertArrayToObjectString(sheetValues) {
    const rows = [...sheetValues]; // deepCopy
    const keys = rows.splice(0, 1)[0];
    return rows.map(row => {
      const obj = {};
      row.map((item, index) => {
        // 保存されている値が0の場合にfalse判定になりnullが出力されてしまうのでString(item) === ""の記述で判定を行っている。
        // obj[String(keys[index])] = String(item) === "" ? null : String(item);
        // すべての値を文字型にして出力する方が関数としての役割。JavaScript上では扱いやすい。
        obj[String(keys[index])] = String(item);
      });
      return obj;
    });
  }

  static deepCopy(val) {
    if (typeof val !== 'object' || val === null) {
      return val;
    } else if (['string', 'boolean', 'number'].includes(typeof val)) {
      return val;
    } else if (Array.isArray(val)) {
      return val.map(item => SheetDB_.deepCopy(item));
    } else if (val instanceof Date) {
      return new Date(val.getTime());
    } else {
      return Object.keys(val).reduce((acc, key) => {
        acc[key] = SheetDB_.deepCopy(val[key]);
        return acc;
      }, {});
    }
  }
}
