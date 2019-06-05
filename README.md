FixFont
=======

FixFontはPowerPointのプレゼンテーション資料のフォントを統一します。

PowerPointでプレゼンテーション資料を作成する際に、他の人が作った資料からスライドをコピーしてきたりすると、フォントがバラバラになってしまうことがあります。

わかりやすい資料を作成するためにフォントを統一したいのですが、以下の理由から容易ではありません。

* スライドマスタのフォントを変更しても、テキストボックス等で追加したものは変更の範囲外
* PowerPointの「フォントの置換」機能は、使用されている全てのフォントについて1つ1つ置換後のフォントの指定が必要
* PowerPointの「フォントの置換」機能に、置換前と置換後のフォントに指定できない組み合わせが存在

FixFontはこの問題を解決します。FixFontはPowerPointのプレゼンテーションのフォントを、そのプレゼンテーションのテーマのフォント(デフォルトのフォント)に修正します。

使い方
------

FixFontコマンドに引数としてPowerPointファイルを指定します。

```console
FixFont PPT_FILE...
```

* FixFontは指定されたファイルを順次バックアップしてから開き、フォントを修正して保存します。(`sample.pptx` は `sample - backup.pptx` にバックアップされる)
* 修正状況は画面に出力されるだけではなく、PowerPointファイルと同名のログ・ファイルに記録されます。(`sample.pptx` のフォント修正は `sample.log` に記録される)
* ログに記録されるテーマのフォントの意味は以下のとおりです。

  フォント | 意味
  ---------|-------------------------
  +mj-lt   | 見出しのフォント(英数字)
  +mn-lt   | 本文のフォント(英数字)
  +mj-ea   | 見出しのフォント(日本語)
  +mn-ea   | 本文のフォント(日本語)

Tips
----

エクスプローラーの［送る］メニューに追加すると便利です。

参考: [［送る］メニューに項目を追加する方法（Windows 7／8.x／10編）：Tech TIPS - ＠IT](https://www.atmarkit.co.jp/ait/articles/1109/30/news131.html)

Author
------

* [Mikio Ogawa](https://github.com/miogawa)
* [Shinichi Akiyama](https://github.com/shakiyam)

License
-------

[MIT License](https://opensource.org/licenses/mit)
