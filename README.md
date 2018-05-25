# Generating the data for a language recognition model(based on CNNs).

This is a part of the data collection/generation work I did at Information Sciences Institute, Marina del Ray.

## Components of the code:
1. Google custom Search.
2. Win32com package to work with the powerpoint slides using python
3. PIL (pillow) to save the exported image.

## Why is this useful?
This code works with the win32com library.
It is difficult to work with this library:
1. It only works with Windows Machine.
2. The documentation you have to refer to is developed for c# and makes it very frustrating to 'guess' the right method or property.

In addition to this is a cheap way to build a data set for a script recognition system based on Neural Networks. The data collected from this method will atleast allow you to train a model which can recognize the script of the language in a given powerpoint.

## What does this code do:
Here are the tasks this code carries out:
1. Retrieve the links of the powerpoints from the web using Google Custom Search.
2. Download the powerpoints from links in the previous step.
3. Works on these powerpoints to extract the images in 720 x 540 size. 
4. It also generates an annotated file with the coordinates if the text found on the slide.

## Steps to run this code.

1. Getting the powerpoint slides from the web using custom search.
    1. open up w_web_interactions.py and paste the key in api_key and the search_engine_id. Please follow [follow me to get the key](https://stackoverflow.com/questions/37083058/programmatically-searching-google-in-python-using-custom-search)
    2. 
