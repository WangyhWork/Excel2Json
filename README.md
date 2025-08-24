# Excel2Json

**Excel2Json** は **Excel VSTO アドイン** で、Excel ワークブックのデータを自動的に **JSON 形式** にエクスポートするツールです。  
Excel の **保存** や **閉じる** 操作に連動して JSON を生成し、さらに **参照リスト** や **型定義**、ドロップダウンによるデータ検証にも対応しています。

## 主な機能
- Excel の保存/終了時に JSON を自動エクスポート  
- `ClosedXML`、`NPOI`、`Newtonsoft.Json` を利用した高速かつ安定した処理  
- 参照リストを利用したデータ参照に対応  
- Ribbon UI から **Enable Export** チェックボックスでエクスポートの有効/無効を切り替え可能  
- CRC32 計算やサブクラス対応によるネストされた JSON 構造の出力 


## 動作環境
- Windows 10 / 11  
- Microsoft Excel 2016 以降  
- .NET Framework 4.7.2+  
- Visual Studio（VSTO アドインのビルド用）  
- NuGet パッケージ  
  - `ClosedXML`  
  - `NPOI`  
  - `Newtonsoft.Json`


## インストール手順
1. [Releases](https://github.com/WangyhWork/Excel2Json/releases) から **「Excel2Json.zip」** をダウンロードし、解凍します。  
   解凍すると **「Excel2Json」** フォルダが表示され、その中に **Setup.exe** があります。  
   `Setup.exe` をダブルクリックしてインストールを開始してください。  

2. インストーラーメッセージが表示されたら、**インストール** をクリックします。  

3. インストール中に **SecurityException** が出る場合の対処法：  
   - **インストールエラー1**  
     - `Excel2Json` フォルダ内の `Excel2Json.vsto` を右クリック → プロパティ → 「ブロックの解除」を実行してください。  

   - **インストールエラー2**  
     - `Excel2Json.vsto` を右クリックして「ブロックの解除」  
     - `Excel2Json\Application Files\Excel2Json_1_0_0_2\Excel2Json.dll.manifest` も「ブロックの解除」  
     - これでインストール可能になります。  

   - **インストールエラー3**  
     - 一度アンインストールして再インストールすると解消することがあります。  

4. インストール完了後、Excel を開くと **「アドイン」タブ** が追加されます。  
   `Enable Export` にチェックを入れると最初の JSON 出力が実行され、その後は保存や終了時に自動でエクスポートが行われます。  


## アンインストール（削除）手順
Windows の設定 → 「アプリ」 → インストール済みアプリ一覧で **「Json」** と検索すると、**「Excel2Json」** アプリが表示されます。  
そこから削除を実行してください。  

 
## 使い方
1. Visual Studio でソリューションをビルドし、Excel にアドインをインストールします。  
2. Excel を起動すると、新しいリボンタブ **「アドイン → Excel2Json」** が追加されます。  
3. **Enable Export** を有効化すると、Excel ファイルを保存または閉じたときに JSON が自動出力されます。  
4. 出力される JSON の保存先は元の Excel ファイルと同じディレクトリに保存されます。
