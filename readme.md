# ExcelChatGPT

ワークシート関数からChatGPTを呼び出して回答をセルに出力します。



## 設定

設定シートに必要情報を入力

![image](https://github.com/ikuma-hiroyuki/ExcelChatGPT/blob/master/sample0.png)


## ワークシート関数としての ChatGPT関数

```=ChatGPT(userContent1, [userContent2], [output], [role])```

ワークシート関数として呼び出すとユーザーがセルに何か入力するたびに全関数が再計算され長大な待ち時間が発生します。



## マクロから ChatGPT関数を実行

Alt+F8キーでマクロを呼び出し、ExecuteChatGPTを実行すると選択範囲にのみ応答を出力します。

関数に与える引数は下図のように設定します。

![image](https://github.com/ikuma-hiroyuki/ExcelChatGPT/blob/master/sample1.png)
