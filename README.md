```
python test_excel.py test.xlsx
```

```
----- xlrd==1.2.0 result -----

[['ID', 'Val1', 'Val2', 'Val3', 'Val4', 'Val5', 'Val6'],
 ['ID1', 1.0, '1', '1.0', '00001111', 'A', 1.1],
 ['ID2', 2.0, '2', '2.0', '00002222', 'B', 2.2],
 ['ID3', 3.0, '3', '3.0', '00003333', 'C', 3.3],
 ['ID4', 4.0, '4', '4.0', '00004444', 'D', 4.4],
 ['ID5', 5.0, '5', '5.0', '00005555', 'E', 5.5],
 ['', '', '', '', '', '', ''],
 ['ID6', 6.0, '6', '6.0', '00006666', 'F', 6.6]]

----- pylightxl result -----

[['ID', 'Val1', 'Val2', 'Val3', 'Val4', 'Val5', 'Val6'],
 ['ID1', 1.0, '1', '1.0', '00001111', 'A', 1.1],
 ['ID2', 2.0, '2', '2.0', '00002222', 'B', 2.2],
 ['ID3', 3.0, '3', '3.0', '00003333', 'C', 3.3],
 ['ID4', 4.0, '4', '4.0', '00004444', 'D', 4.4],
 ['ID5', 5.0, '5', '5.0', '00005555', 'E', 5.5],
 ['', '', '', '', '', '', ''],
 ['ID6', 6.0, '6', '6.0', '00006666', 'F', 6.6]]

----- pandas result -----

ID       object
Val1    float64
Val2    float64
Val3    float64
Val4    float64
Val5     object
Val6    float64
dtype: object
[['ID', 'Val1', 'Val2', 'Val3', 'Val4', 'Val5', 'Val6'],
 ['ID1', 1.0, 1.0, 1.0, 1111.0, 'A', 1.1],
 ['ID2', 2.0, 2.0, 2.0, 2222.0, 'B', 2.2],
 ['ID3', 3.0, 3.0, 3.0, 3333.0, 'C', 3.3],
 ['ID4', 4.0, 4.0, 4.0, 4444.0, 'D', 4.4],
 ['ID5', 5.0, 5.0, 5.0, 5555.0, 'E', 5.5],
 [nan, nan, nan, nan, nan, nan, nan],
 ['ID6', 6.0, 6.0, 6.0, 6666.0, 'F', 6.6]]

----- openpyxl result -----

[('ID',
  'Val1',
  'Val2',
  'Val3',
  'Val4',
  'Val5',
  'Val6',
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None,
  None),
 ('ID1', 1.0, '1', '1.0', '00001111', 'A', 1.1),
 ('ID2', 2.0, '2', '2.0', '00002222', 'B', 2.2),
 ('ID3', 3.0, '3', '3.0', '00003333', 'C', 3.3),
 ('ID4', 4.0, '4', '4.0', '00004444', 'D', 4.4),
 ('ID5', 5.0, '5', '5.0', '00005555', 'E', 5.5),
 (None, None, None, None, None),
 ('ID6', 6.0, '6', '6.0', '00006666', 'F', 6.6),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None),
 (None, None, None, None, None)]
```


openpyxl what is wrong with you?