# 各建筑坐标存储在表格中，打印完路线后出现地图，地图上红线为应该走的路线，注意两个地图文件的存储位置和文件类型

# 调用pillow库，matplotlib库，numpy库三个库，前两个用于绘制图片
from PIL import Image
from pylab import *
import numpy as np
import logging
from pywebio.input import *

# 注意两个地图文件存储位置

map_path = '..\\地图文件\\map_school.txt'
scmap_path = '..\\地图文件\\scmap.jpg'

# 用于将二维数组中存储的坐标与图片的坐标进行互换
x_dict = {0: 43, 1: 90, 2: 127, 3: 158, 4: 200, 5: 240, 6: 280, 7: 320, 8: 345, 9: 375, 10: 406, 11: 440, 12: 475,
          13: 505, 14: 540, 15: 565 \
    , 16: 617, 17: 645, 18: 680, 19: 715, 20: 742, 21: 778, 22: 806, 23: 858, 24: 909, 25: 950, 26: 985, 27: 1017,
          28: 1043, 29: 1085, 30: 1125}
y_dict = {0: 80, 1: 93, 2: 129, 3: 148, 4: 185, 5: 219, 6: 257, 7: 285, 8: 320, 9: 350, 10: 388, 11: 413, 12: 440,
          13: 480, 14: 508, 15: 528 \
    , 16: 560, 17: 585, 18: 635, 19: 675, 20: 750, 21: 800, 22: 845, 23: 892, 24: 930, 25: 984, 26: 1033, 27: 1100}

pre_route = list()  # 宽度搜索得到的节点
q = list()  # 队列结构控制循环次数
xchange = [0, 1, 0, -1]  # 节点的移动
ychange = [1, 0, -1, 0]
visited = list()  # 记录节点是否已遍历
father = list()  # 每一个pre_route节点的父节点
route = list()
way = list()  # 用于存储需要经历的节点
allroute = list()


def clearList():
    father.clear()
    pre_route.clear()
    route.clear()
    q.clear()


def inputway():
    pre = 0
    point = list(input("输入节点（以空格为间隔，第一个为当前位置，后续为若干按序需要到达的地点）：").split(' '))
    for i in range(len(point)):
        point[i] = int(point[i])
    while pre < len(point):
        now = np.array((point[pre], point[pre + 1]))
        way.append(now)
        pre = pre + 2

    return way


def mapcreate():
    # 读取文件
    a = np.loadtxt(map_path, dtype='int')

    # 确定地图最大规格
    mx = np.max(a) + 1

    # 地图创建
    map = np.ones((mx, mx), dtype='int')

    # 地图标点
    for i in range(a.shape[0]):
        if a[i][0] == a[i][2]:
            if a[i][1] < a[i][3]:
                np.put(map[a[i][0]], [range(a[i][1], a[i][3] + 1)], 0)
            else:
                np.put(map[a[i][0]], [range(a[i][3], a[i][1] + 1)], 0)
        if a[i][1] == a[i][3]:
            if a[i][0] < a[i][2]:
                for j in range(a[i][0], a[i][2] + 1):
                    map[j][a[i][1]] = 0
            else:
                for j in range(a[i][2], a[i][0] + 1):
                    map[j][a[i][1]] = 0
    return map


def searchway(l, x, y, m, n):
    visited = [[0 for i in range(len(l[0]))] for j in range(len(l))]

    # 入口节点设置为已遍历
    visited[x][y] = 1
    q.append([x, y])

    # 队列为空则结束循环
    while q:
        now = q[0]
        # 移除队列头结点
        q.pop(0)
        for i in range(4):
            # 当前节点
            point = [now[0] + xchange[i], now[1] + ychange[i]]
            if point[0] < 0 or point[1] < 0 or point[0] >= len(l) or point[1] >= len(l[0]) or visited[point[0]][
                point[1]] == 1 or l[point[0]][point[1]] == 1:
                continue
            father.append(now)
            visited[point[0]][point[1]] = 1
            q.append(point)
            pre_route.append(point)
            if point[0] == m and point[1] == n:
                return 1
    # print("[%d,%d]节点无到达路线，结束" % (m, n))
    return 0


# 输出最短路径
def get_route(father, pre_route):
    route = [pre_route[-1], father[-1]]
    for i in range(len(pre_route) - 1, -1, -1):
        if pre_route[i] == route[-1]:
            route.append(father[i])
    route.reverse()
    # print("最短路线为：\n", route)
    # print("距离为：", len(route) - 1)
    return route


# if __name__ =='__main__':
def map_init(way, name):
    way.insert(0,[5,3])
    allroute = []#新添加
    map = mapcreate()
    x_point = list()
    y_point = list()
    for n in range(len(way)):
        x_point.append(x_dict[way[n][1]])
        y_point.append(y_dict[way[n][0]])

    i, j = 0, 1
    while j != len(way):
        if searchway(map, way[i][0], way[i][1], way[j][0], way[j][1]) == 1:
            route = get_route(father, pre_route)
            allroute = allroute + route[1:]
            clearList()
            i, j = i + 1, j + 1
        else:
            break
    # print("路线导航结束")
    # print("接下来的路线为：", allroute)

    route_len = len(allroute)
    x_list = list()
    y_list = list()

    for m in range(route_len):
        x_list.append(x_dict[allroute[m][1]])
        y_list.append(y_dict[allroute[m][0]])

    #此处暂时关闭日志记录，防止图像相关内容记录到日志中
    logging.disable(logging.CRITICAL)

    # 读取图像到数组中
    im = array(Image.open(scmap_path))
    figure(num=name)

    # 绘制有坐标轴的
    subplot(111)
    imshow(im)

    plot(x_point, y_point, '*')
    plot(x_list[:route_len], y_list[:route_len], 'r')
    show()

    #重启日志记录
    logging.disable(logging.NOTSET)
