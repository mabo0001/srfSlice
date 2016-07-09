#画出的polyline不要旋转

import win32com
import os
import os.path
import numpy as np
import numbers
import array
from win32com.client import Dispatch, constants

def srfSiliceOne(linenum,srffilename,gridfilename,xMapPerPU=20,yMapPerPU=20):
    srf = win32com.client.DispatchEx ('Surfer.Application')
    srf.visible = 1

    srfdoc = srf.documents.open (srffilename)
    srfshapes = srfdoc.shapes
    for srfshape in srfshapes:
        if ("Map" == srfshape.name):
            mapframe = srfshape
        if ("Polyline" == srfshape.name):
           polyline = srfshape

    print (mapframe.name)

    ##获取polyline的坐标
    vertices = polyline.Vertices
    basepolyline = srfdoc.shapes.AddPolyLine (vertices)
    basepolyline.line.width = polyline.line.width

    polylinex = list (polyline.Vertices[::2])
    polyliney = list (polyline.Vertices[1::2])

    polylinex = [x - (basepolyline.left - polyline.left) for x in polylinex]
    polyliney = [y - (basepolyline.top - polyline.top) for y in polyliney]
    srfdoc.shapes.AddSymbol (polylinex[0], polyliney[0])
    ##查找contourmap的坐标
    contourmap = mapframe.overlays.Item (1)
    ##topaxis = mapframe.axes.Item(2)
    bottomaxis = mapframe.axes.Item (1)
    leftaxis = mapframe.axes.Item (3)
    grid_xmin = mapframe.xMin
    grid_ymin = mapframe.yMin
    xmin = leftaxis.left + leftaxis.width
    ymin = bottomaxis.top
    ##
    ##查找contourmap对应的grid文件，设置grid文件属性跟srf文件属性一致
    gridfile = gridfilename
    ##contourmap.gridfile = gridfile
    ##grid = contourmap.grid
    mapframe.xMapPerPU = 20
    mapframe.yMapPerPU = 20
    srfdoc.shapes.AddSymbol (xmin, ymin)
    grid_data_x = [(plx - xmin) * mapframe.xMapPerPU + grid_xmin for plx in polylinex]
    grid_data_y = [(ply - ymin) * mapframe.yMapPerPU + grid_ymin for ply in polyliney]
##
    ##通过线性插值，得到polyline的实际坐标

    count = int ((mapframe.xMax - mapframe.xMin) / 10 + 1)
    xi = [i + mapframe.xMin for i in range (0, count * 10, 10)]
    yi = np.interp (xi, grid_data_x, grid_data_y)
    ###往上20m
    yi = yi + -20
    ###往下20m
    #yi = yi - 20
    f = open ("slice.bln", "w+")
    f.writelines (len (yi).__str__ () + " " + "0" + "\n")
    for i in range (len (xi)):
        f.writelines (xi[i].__str__ () + " " + yi[i].__str__ () + "\n")
    f.close ()
##通过srf.slice进行切片
    srf.gridslice (gridfile, r"c:\Users\lenovo\PycharmProjects\mabo0001\slice.bln", "",
               r"c:\Users\lenovo\PycharmProjects\mabo0001\out.txt", 0, 0)

##读取切片后的数据
###out = np.loadtxt(r"c:\Users\lenovo\PycharmProjects\mabo0001\out.dat",dtype=np.str,delimiter=" ")
# print(out)
    b = np.loadtxt (r"out.txt", dtype=np.float64)
    b[:, 4] = linenum
    c = b[np.in1d (b[:, 0], xi), :]
    out = c[:, [-1, 0, 2]]
    srfdoc.close (2,srfdoc.FullName)
    srf.quit ()
    return out

outdata=np.arange(3).reshape(1,3)
for root, dirs, files in os.walk (r'C:\Users\lenovo\Desktop\中阳瞬变\2区'):
    #print("==outdata=========================")
    #print(outdata)
    for file in files:
        if (file[-4:] == ".srf"):
            #print (os.path.join (root, file))
            #print(root)
            #print(file)
            srffilename = os.path.join (root, file)
            print(srffilename)
            gridfilename = os.path.join (root,"out1.grd")
            print(gridfilename)
            strlinenum="".join(list(filter (str.isdigit,file)))
            out = srfSiliceOne(np.float64(strlinenum),srffilename,gridfilename,20,20)
            outdata = np.vstack((outdata,out))
print("outdata+++++++++++++++++++++++++++++++++++++++++")
print(outdata)
np.savetxt(r'c:\Users\lenovo\PycharmProjects\mabo0001\data.dat',outdata,fmt="%5.2f")

