# Table_extraction
Extract the table structure and cell contents from the image, and generate the corresponding Excel file.

## Extracted Results
### ·table1:
![1](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/8c8845c4-138a-463a-a09f-0dbdcb72790c)

### ·extract1.xlsx:
![image](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/51f54a15-33f5-436a-9385-d3a02afa94de)

### ·table2:
![2](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/abc512c6-6877-4a67-ac9b-ebb32a40aee5)
### ·extract2.xlsx：
![image](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/1dd3b45b-bf39-40ff-8875-002205cfd6ab)

## Environment
```python
import cv2
import numpy as np
import xlsxwriter
import pytesseract
```
## Unresolved Issues
Inaccurate text recognition in some cells.
