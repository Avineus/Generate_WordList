# Generate_WordList
To generate list of all word list with part of speech such as CVC, CCVC etc..

This project is to list all possible CVC, CCVCC, and possible combination of words. Mainly this is written to create flash cards for kids with all possible combination of words.

This project also verify word created with dictionary and gets the adverbs, noun and all parts of speech

Script is enhancement to write the words in a Word document as table.

Pre-request to run this script is to install the below python modules as shown below

Download free dictionary from below link https://freedict.org/freedict-database.json
Install below python pip

apt-get install python3-pip pip3 install pyenchant pip3 install nltk pip3 install python-docx

Open python3 and type below command

import nltk nltk.download('punkt') nltk.download('averaged_perceptron_tagger') nltk.download('universal_tagset') nltk.download('wordnet')
