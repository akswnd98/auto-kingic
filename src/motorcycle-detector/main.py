import os
from time import time
from darknet.python import darknet as dn
from ctypes import *

meta = dn.load_meta(b'./darknet/cfg/coco.data')
net = dn.load_net(b'darknet/cfg/yolov3.cfg', b'./darknet/weights/yolov3.weights', 0)

f = open('result.txt', 'w')

files = os.listdir('violation-images')
for file in files:
  detections = dn.detect(net, meta, os.path.join('violation-images', file).encode('ascii'))
  motorbikes = list(filter(lambda detection: detection[0].decode() == 'motorbike', detections))
  if len(motorbikes) > 0:
    f.write('{}\n'.format(file))

f.close()
