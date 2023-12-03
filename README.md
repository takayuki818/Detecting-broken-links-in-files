# Name
リンク切れ検査ツール.xlsm
## Overview
ファイルを開いた際に、  
![リンクが更新できません](https://github.com/takayuki818/Detecting-broken-links-in-files/assets/147408435/8160b35f-aa84-4114-bc6b-7e5c2f6f5d9a)  
![外部ソースへのリンクが含まれています](https://github.com/takayuki818/Detecting-broken-links-in-files/assets/147408435/25fa5127-f64d-4d7c-9a66-55397fb0fdd4)  
↑の様なメッセージが出てしまうExcelファイルの問題発生箇所を探査します。
## Usage
1.「MENU」シートの「Excelファイル選択→リンク切れ検査」ボタンを押下  
2.ダイアログボックスから検査対象のExcelファイルを選択  
　→　検査結果メッセージを返します。
## Description
### メッセージ内容説明  
・名前の定義のリンク切れ箇所候補：  
　「数式」タブの「名前の管理」に設定している定義名について、リンク切れや参照エラーが生じている可能性があります。  
・リスト型入力規則のリンク切れ箇所候補：  
　「データ」タブの「データの入力規則」を設定しているセルについて、リンク切れや参照エラーが生じている可能性があります。
