# Table_extraction
Extract the table structure and cell contents from the image, and generate the corresponding Excel file.

## Extracted Results
·table1:
![1](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/8c8845c4-138a-463a-a09f-0dbdcb72790c)
·extract1.xlsx:
![image](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/e03683ad-48cd-4f18-bd54-d3e36a638a4b)

·table2:
![2](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/abc512c6-6877-4a67-ac9b-ebb32a40aee5)
·extract2.xlsx：
![image](https://github.com/Pyxis-xu/Table_extraction/assets/130300323/3ba44e80-082e-4e94-9c08-99cd857aa5b0)

## Environment
```python
import cv2
import numpy as np
import xlsxwriter
import pytesseract
```
## Unresolved Issues
Inaccurate text recognition in some cells.
