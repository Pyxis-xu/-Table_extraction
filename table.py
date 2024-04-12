import cv2
import numpy as np
import xlsxwriter

src = '1.png'
raw = cv2.imread(src, 1)

gray = cv2.cvtColor(raw, cv2.COLOR_BGR2GRAY)
binary = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, -5)
rows, cols = binary.shape
scale2 = 15
scale = 20

kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (cols // scale, 1))
kernel1 = cv2.getStructuringElement(cv2.MORPH_RECT, (cols // scale2, 1))
eroded = cv2.erode(binary, kernel, iterations=1)
dilated_col = cv2.dilate(eroded, kernel1, iterations=1)
kernel2 = cv2.getStructuringElement(cv2.MORPH_RECT, (1, rows // scale2))
kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, rows // scale))
eroded = cv2.erode(binary, kernel, iterations=1)
dilated_row = cv2.dilate(eroded, kernel2, iterations=1)

bitwise_and = cv2.bitwise_and(dilated_col, dilated_row)
merge = cv2.add(dilated_col, dilated_row)
ret, binary = cv2.threshold(merge, 127, 255, cv2.THRESH_BINARY)
contours, _ = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
area = []
for k in range(len(contours)):
    area.append(cv2.contourArea(contours[k]))
max_idx = np.argmax(np.array(area))
m_d_r = []
m_u_l = []
max_p = 0
min_p = 1e6
for l1 in contours[max_idx]:
    for l2 in l1:
        if sum(l2) > max_p:
            max_p = sum(l2)
            d_r = l2
        if sum(l2) < min_p:
            min_p = sum(l2)
            u_l = l2
m_d_r = d_r
m_u_l = u_l
padding = 5
x0 = max(m_u_l[0] - padding, 0)
x1 = min(m_d_r[0] + padding, raw.shape[1])
y0 = max(m_u_l[1] - padding, 0)
y1 = min(m_d_r[1] + padding, raw.shape[0])

bitwise_and_crop = bitwise_and[y0:y1, x0:x1]
merge = merge[y0:y1, x0:x1]
raw = raw[y0:y1, x0:x1]

ys, xs = np.where(bitwise_and_crop > 0)
y_point_arr = []
x_point_arr = []
i = 0
sort_x_point = np.sort(xs)
for i in range(len(sort_x_point) - 1):
    if sort_x_point[i + 1] - sort_x_point[i] > 3:
        x_point_arr.append(sort_x_point[i])
        i += 1
x_point_arr.append(sort_x_point[i])
i = 0
sort_y_point = np.sort(ys)
for i in range(len(sort_y_point) - 1):
    if sort_y_point[i + 1] - sort_y_point[i] > 3:
        y_point_arr.append(sort_y_point[i])
        i += 1
y_point_arr.append(sort_y_point[i])

h_list = [y_point_arr[i + 1] - y_point_arr[i] for i in range(len(y_point_arr) - 1)]
w_list = [x_point_arr[i + 1] - x_point_arr[i] for i in range(len(x_point_arr) - 1)]

col_alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

workbook = xlsxwriter.Workbook('extract.xlsx')
worksheet = workbook.add_worksheet()

for i in range(len(w_list)):
    worksheet.set_column('{}:{}'.format(col_alpha[i], col_alpha[i]), w_list[i] / 6)
for j in range(len(h_list)):
    worksheet.set_row(j, h_list[j])

def islianjie(p1, p2, img):
    if p1[0] == p2[0]:
        for i in range(min(p1[1], p2[1]), max(p1[1], p2[1]) + 1):
            if sum([img[j, i] for j in range(max(p1[0] - 5, 0), min(p1[0] + 5, img.shape[0]))]) == 0:
                return False
        return True
    elif p1[1] == p2[1]:
        for i in range(min(p1[0], p2[0]), max(p1[0], p2[0]) + 1):
            if sum([img[i, j] for j in range(max(p1[1] - 5, 0), min(p1[1] + 5, img.shape[1]))]) == 0:
                return False
        return True
    else:
        return False

class cell:
    def __init__(self, lt, rd, belong):
        self.lt = lt
        self.rd = rd
        self.belong = belong

lt_list_x = x_point_arr[:-1]
lt_list_y = y_point_arr[:-1]
rd_list_x = x_point_arr[1:]
rd_list_y = y_point_arr[1:]

d = {}
for i in range(len(lt_list_x)):
    for j in range(len(lt_list_y)):
        d['cell{}_{}'.format(i, j)] = cell([lt_list_x[i], lt_list_y[j]], [rd_list_x[i], rd_list_y[j]], [lt_list_x[i], lt_list_y[j]])

for i in range(len(lt_list_x)):
    for j in range(len(lt_list_y)):
        p1 = [d['cell{}_{}'.format(i, j)].rd[1], d['cell{}_{}'.format(i, j)].lt[0]]
# 左下点
        p2 = [d['cell{}_{}'.format(i, j)].rd[1], d['cell{}_{}'.format(i, j)].rd[0]]  # 右下点
        p3 = [d['cell{}_{}'.format(i, j)].lt[1], d['cell{}_{}'.format(i, j)].rd[0]]  # 右上点
        if not islianjie(p1, p2, merge):
            d['cell{}_{}'.format(i, j + 1)].belong = d['cell{}_{}'.format(i, j)].belong
        if not islianjie(p2, p3, merge):
            d['cell{}_{}'.format(i + 1, j)].belong = d['cell{}_{}'.format(i, j)].belong

crop_list = {}
for i in range(len(lt_list_x)):
    for j in range(len(lt_list_y)):
        crop_list['{},{}'.format(d['cell{}_{}'.format(i, j)].belong[0], d['cell{}_{}'.format(i, j)].belong[1])] = d[ 'cell{}_{}'.format(i, j)].rd

w_h_list = []
zmax = 0
zmin = 1e6
zlt = []
zrd = []

for key in crop_list.keys():
    lt = [int(i) for i in key.split(',')]
    rd = crop_list[key]
    if sum(rd) > zmax:
        zrd = rd
        zmax = sum(rd)
    if sum(lt) < zmin:
        zlt = lt
        zmin = sum(lt)
    cv2.imwrite('crop/{}.jpg'.format(key), raw[lt[1]:rd[1], lt[0]:rd[0]])

merge_format = workbook.add_format({ 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'border_color': '000000' })

for key in crop_list.keys():
    lt = [int(i) for i in key.split(',')]
    rd = crop_list[key]
    lt_ = [lt[0] - zlt[0], lt[1] - zlt[1]]
    rd_ = [rd[0] - zlt[0], rd[1] - zlt[1]]

    for i in range(len(w_list) + 1):
        if lt_[0] == sum(w_list[:i]):
            lt_col = chr(ord('A') + i)
        if rd_[0] == sum(w_list[:i]):
            rd_col = chr(ord('A') + i - 1)
    for i in range(len(h_list) + 1):
        if lt_[1] == sum(h_list[:i]):
            lt_row = i + 1
        if rd_[1] == sum(h_list[:i]):
            rd_row = i
    if lt_col == rd_col and lt_row == rd_row:
        worksheet.write('{}{}'.format(lt_col, lt_row), '', merge_format)
    else:
        worksheet.merge_range('{}{}:{}{}'.format(lt_col, lt_row, rd_col, rd_row), '', merge_format)

workbook.close()
