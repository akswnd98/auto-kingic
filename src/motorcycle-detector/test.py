import os
from time import time
from darknet.python import darknet as dn
from ctypes import *

meta = dn.load_meta(b'./darknet/cfg/coco.data')
net = dn.load_net(b'darknet/cfg/yolov3.cfg', b'./darknet/weights/yolov3.weights', 0)

cur = time()
print(cur)
print(dn.detect(net, meta, b'./a.jpg'))
print(dn.detect(net, meta, b'./b.jpg'))
print(time() - cur)
