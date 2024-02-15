# README

分析化学において使用する数値丸め用のExcelマクロ関数の定義について記録します。

# 使い方
- Round4AnalyticalChemistry.basをExcelのVBEditor画面でインポートすると使用できます
- Excelのワークシート関数として使用できるため、他の関数のようにセルに `=function(...)` と入力することで使用できます。

# 関数の使い方

## round2
この関数では、 `digits` で指定した有効数字桁数で四捨五入を行います。

### 引数
```vb
round2(x, digits)

- x: 丸めを適用する数字
- digits: 有効数字桁数
```

### 使用例

```vb
v = Array(0.001234, 0.02011, 0.20113, 9.801, 123.52)
For Each num In v
    Debug.Print num & " = " & round2(num, 2)
Next
```

```vb
0.001234 = 0.0012
0.02011 = 0.02
0.20113 = 0.2
9.801 = 9.8
123.52 = 120
```

## roundJIS
この関数では、 `digits` で指定した有効数字桁数で **JIS Z 8401** を適用します。 JIS Z 8401 の詳細はJISのページをご確認ください。

### 引数
```vb
roundJIS(x, digits)

- x: 丸めを適用する数字
- digits: 有効数字桁数
```

### 使用例

```vb
v = Array(0.00122501, 0.001225, 0.001235, 12.35, 12450)
For Each num In v
    Debug.Print num & " = " & roundJIS(num, 3)
Next
```

```vb
0.00122501 = 0.00123
0.001225 = 0.00122
0.001235 = 0.00124
12.35 = 12.4
12450 = 12400
```

## LLQ, LLD
この関数では標準偏差から定量下限値(LLQ), 検出下限値(LLD)を計算し、 `digits` で指定した有効数字桁数で四捨五入を行います。  
LLDはLLQの有効数字桁数に合わせて切り捨て処理がされます。

### 引数
```vb
LLQ(sd, digits)
LLD(sd, digits)

- sd: 標準偏差
- digits: 有効数字桁数
```

### 使用例

```vb
d = Array(0.001, 0.0023, 0.0013, 0.0066, 0.0035)

'' 標準偏差の計算
sd = WorksheetFunction.StDev(d)
Debug.Print "SD: " & sd

'' 定量下限値の計算
Debug.Print "LLQ: " & LLQ(sd)

'' 検出下限値の計算
Debug.Print "LLD: " & LLD(sd)
```

```vb
SD: 2.26781833487605E-03
LLQ: 0.023
LLD: 0.006
```

## roundSet
この関数では、 指定した有効数字桁数で四捨五入を行います。また、参照する数値及び有効数字桁数を指定し、参照値以下の桁は切り捨て処理を行います。

### 引数
```vb
roundSet(x, y, x_digits, y_digits, fitRoundJIS)

- x: 丸めを適用する数字
- y: 参照値
- x_digits: xの有効数字桁数
- y_digits: yの有効数字桁数
- fitRoundJIS: 
    True: JIS Z 8401の適用
    False: 通常の四捨五入を適用
```

### 使用例

```vb
v = Array(0.001234, 0.02011, 0.20113, 9.801, 123.52)
sd = WorksheetFunction.StDev(d)
Debug.Print "LLD: " & LLD(sd)
For Each num In v
    Debug.Print num & " = " & roundSet(num, LLQ(sd), 3, 2, False)
Next
```

```vb
LLD: 0.006
0.001234 = 0.001
0.02011 = 0.020
0.20113 = 0.201
9.801 = 9.80
123.52 = 124
```

## roundSetStyle
この関数では、 `roundSet` と同様の数値丸めに加え、 *検出下限値未満* 及び *検出下限値以上定量下限値未満* の場合に表記を変更することができます。  


適用例: 
*検出下限値未満* : '<[検出下限値]'
*検出下限値以上定量下限値未満* : '(数値)'


### 引数
```vb
roundSetStyle(x, sd, digits, ll_digits, prefix1, suffix1, prefix2, suffix2, fitRoundJIS)

- x: 丸めを適用する数字
- sd: 標準偏差
- digits: xの有効数字桁数
- ll_digits: 定量下限値及び検出下限値の有効数字桁数
- prefix1: 検出下限値未満の際に、数値の頭に付ける記号
- suffix1: 検出下限値未満の際に、数値の末尾に付ける記号
- prefix2: 検出下限値以上定量下限値未満の際に、数値の頭に付ける記号
- suffix2: 検出下限値以上定量下限値未満の際に、数値の末尾に付ける記号
- fitRoundJIS: 
    True: JIS Z 8401の適用
    False: 通常の四捨五入を適用
```

### 使用例

```vb
v = Array(0.001234, 0.02011, 0.20113, 9.801, 123.52)
sd = WorksheetFunction.StDev(d)
Debug.Print "LLD: " & LLD(sd)
For Each num In v
    Debug.Print num & " = " & roundSetStyle(num, sd, 3, 2, "<[", "]", "(", ")", False)
Next
```

```vb
LLD: 0.006
0.001234 = <[0.006]
0.02011 = (0.020)
0.20113 = 0.201
9.801 = 9.80
123.52 = 124
```