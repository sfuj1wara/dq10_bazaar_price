# dq10_bazaar_price

## はじめに

ドラクエ10のバザー相場を取得し、エクセルファイルに出力するプログラム

## Requirement

Windows 10 Pro

Python 3.8.2

openpyxl 3.0.4

selenium 3.141.0

ChromeDriver 81.0.4044.69

## Install

[openpyxl 3.0.4](https://openpyxl.readthedocs.io/en/stable/)

```python
> pip install openpyxl
```

[selenium 3.141.0](https://www.selenium.dev/)

```python
> pip install selenium
```

[ChromeDriver](http://chromedriver.chromium.org/downloads)

自分の使っているchromeのバージョンに対応するWebDriverをダウンロードする

## 使い方

必要なライブラリをインストールし、下記コマンドを実行

```python
> python main.py
```

検索するアイテム名を入力してください：が表示されたら、検索したいアイテム名を入力してEnterを押すと、検索したアイテムの旅人バザー上での価格がエクセルシートになって出力されます。
現在は、旅人バザーの一ページ目のみが対象になっています。
