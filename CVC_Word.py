import json
import enchant
import nltk
import re

from docx import Document
from docx.shared import Inches
from nltk.corpus import wordnet 

c = "bcdfghjklmnpqurstvwxyz"
v = "aeiou"

d = enchant.Dict("en_US")

def_str = ""

document = Document()
document.add_heading('CVC Word List', 0)

p = document.add_paragraph('This book has list of CVC words with its associated part of speech. Table with CVC words respents NOUN as NN, VERB as VB, ADJECTIVE as ADJ, ADVERB as ADV')

### Load dictionary json file ####
with open("dictionary.json") as json_file:
  data = json.load(json_file)


### Loop to iterate for each letter in the vowel ####
### This loop will run 5 times a, e, i, o, u ###
for i in v:
   print("CVC words for " + i + " Vowels")
   CVC_Vowel_str = "CVC words for " + i + " Vowels"
   document.add_paragraph(CVC_Vowel_str, style='Intense Quote')

   ## Initialize Count of Words to 0
   count = 0

   ## Add Table
   table = document.add_table(rows=1, cols=3)
   hdr_cells = table.rows[0].cells
   hdr_cells[0].text = 'CVC Words'
   hdr_cells[1].text = 'Part of Speech'
   hdr_cells[2].text = 'Dictionary Meaning'

   ### This loop is to iterate to get C after vowels
   #### In this case it is a[a-z], e[a-z], etc..
   ### This loop will iterate for 21 consonants
   for j in c:
      ### This loop is to iterate to get C at the start
      ### In this case it is eg) at, am will be iterated to get first letter
      ### [a-z]at, [a-z]am, etc..
      ### This loop will iterate for 21 consonants
      for k in c:
         ### Compare formed 3 letter word to the dictionary
         ## Only if word is available in dictory then print
         w = k + i + j
         ### Convert word to upper case to search in json data
         word = w.upper()
         #if word in data : 
         if d.check(w) : 
            count = count + 1;
            ### Dump 3 Letter word and meaning of the word
            ### used utf-8 encoding to avoid UnicodeEncode error 
            #print(word + " " +  data[word].encode('utf-8'))

            token = nltk.word_tokenize(w)
            result = nltk.pos_tag(token, tagset='universal')
            ### To split the word and tag 
            word_tup, tag = zip(*result)

            row_cells = table.add_row().cells
            row_cells[0].text = str(word)

            n_len = len(wordnet.synsets(w, pos='n'))
            v_len = len(wordnet.synsets(w, pos='v'))
            a_len = len(wordnet.synsets(w, pos='a'))
            r_len = len(wordnet.synsets(w, pos='r'))
            pos_str = "NN=" + str(n_len) + " VB=" + str(v_len) + " ADJ=" + str(a_len) + " ADV=" + str(r_len) 
            print(w + " " +  pos_str)
            row_cells[1].text = str(pos_str)

            def_str = "NOUNS \n"
            for noun in wordnet.synsets(w, pos='n'):
                noun_str = noun.definition()
                def_str = def_str + "::" + noun_str + "::\n"

            def_str = def_str + "\n VERBS \n"
            for verb in wordnet.synsets(w, pos='v'):
                verb_str = verb.definition()
                def_str = def_str + "::" + verb_str + "::\n"

            def_str = def_str + "\n ADJECTIVES \n"
            for adjective in wordnet.synsets(w, pos='a'):
                adjective_str = adjective.definition()
                def_str = def_str + "::" + adjective_str + "::\n"

            def_str = def_str + "\n ADVERBS \n"
            for adverb in wordnet.synsets(w, pos='r'):
                adverb_str = adverb.definition()
                def_str = def_str + "::" + adverb_str + "::\n"

            row_cells[2].text = str(def_str)
            print (w + " " + def_str)
            def_str = ""

            document.save('Output/CVC_Word.docx')
            
            

