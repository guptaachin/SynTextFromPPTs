%matplotlib inline
import cv2
import os
import numpy as np 
import sys
from matplotlib import pyplot
import json
import lmdb
from tensorflow.python.client import device_lib
print device_lib.list_local_devices()
from keras.layers import merge, Input, Conv2D, MaxPooling2D, BatchNormalization, Activation, Concatenate
from keras.layers.core import Lambda
from keras.models import Model
from keras import backend as K
from keras.engine.topology import Layer
import numpy as np
import cv2
import urllib, cStringIO
from PIL import Image
from keras.applications.vgg16 import preprocess_input
import tensorflow as tf
from skimage.morphology import label as bwlabel
from skimage.morphology import dilation
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import keras
print keras.__version__
import tensorflow as tf
print tf.__version__
print keras.__file__
data_dir = '/nas/vista-ssd01/users/achingup/imagecollection/images/0_1/'
image_files = [ os.path.join( data_dir, f) for f in filter( lambda f : f.endswith('jpg'), os.listdir( data_dir ) ) ]
print len( image_files )


repo_root = '/nas/vista-ssd01/users/achingup/TDA/TextDetWithScriptID/'#os.path.join( os.getcwd(), os.path.pardir )

print "git repo root =", repo_root
assert os.path.isdir( repo_root ), "ERROR: can't locate git repo for text detection"
model_dir = os.path.join( repo_root, 'model' )
scriptID_weight = os.path.join( model_dir, 'sciptIDModel.h5' )
assert os.path.isfile( scriptID_weight ), "ERROR: can't locate script-ID classification model"
textDet_weight  = os.path.join( model_dir, 'textDetSceneModel.h5' )
assert os.path.isfile( textDet_weight ), "ERROR: can't locate text detection model"

data_dir  = os.path.join( repo_root, 'data' )
lib_dir   = os.path.join( repo_root, 'lib' )
sys.path.insert( 0, lib_dir )

import textDetCore
textDet_model = textDetCore.create_textDet_model()
textDet_model.load_weights( textDet_weight )


def predict_on_rex_models(img_file, viz):    
    import cv2
    import urllib, cStringIO
    from PIL import Image
    from keras.applications.vgg16 import preprocess_input
    img = cv2.imread( img_file )[...,::-1]
    if viz:
        pyplot.figure(figsize=(10,10))
        pyplot.imshow( img )
    x = np.float32( img )
    x = np.expand_dims( x, axis = 0 )
    x = preprocess_input( x )
    proba = textDet_model.predict(x)
    pyplot.figure(figsize=(10,10))
    pyplot.imshow( proba[0])
    pyplot.pause(1)

    model0_mask = (1-proba[0][...,0] > .5).astype('int') 
    return model0_mask


import shutil

def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

dir_unfiltered_images = '/nas/vista-ssd01/users/achingup/imagecollection/images/'
tgt_filtered_path = '/nas/vista-ssd01/users/achingup/TDA/data/filtered_flickr_data/'

dir_list_inf = os.listdir(dir_unfiltered_images)

for each_directory in dir_list_inf:
    dir_unf_full_path = os.join.path(dir_unfiltered_images, each_directory)
    create_directory(os.join.path(tgt_filtered_path, each_directory))
    dir_filtered_path = os.join.path(tgt_filtered_path, each_directory)
    for each_file in os.listdir(dir_unf_full_path):
        img_full_path = os.join.path(dir_unf_full_path, each_file)
        tgt_img_path = os.join.path(dir_filtered_path, each_file)
        # for already processed images
        if (os.path.exists(tgt_img_path)):
            continue
        model0_mask = predict_on_rex_models(img_file, False)
        sum_ = sum(model0_mask.ravel())

        if(sum_ <= 20):
            shutil.copy(img_full_path, tgt_img_path)



