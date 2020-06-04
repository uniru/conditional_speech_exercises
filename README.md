# conditional_speech_exercises
Generate different grammar conditional speech exercises from a given word (.docx) file as input. 

# How to use it
step 1: if you do not have python installed, install it. We reccomend the installation of Anaconda python.

step 2: install the needed addictional python libraries: docx2txt, docx and tkinter (pip install needed_library, es. pip install tkinter)

step 3: open an Anaconda prompt or a terminal, navigate to the location of the python file and type: python conditional_speech_exercises.py

step 4: a window will be opened asking to choose the word file you want to use for the exercise

step 5: it will be asked to choose from the available exercises, just type the number related to the exercise (it will be proposed in the terminal) and press enter

step 6: depending from the exercise choosen the program may ask to select other excel files. 

step 7: the exercise requested has been created with his suffix in the same working directory. 


# Types of exercises availables

1. Preposition substitution: will substitute all preposition of this list ("по", "из", "у", "для", "на", "в", "от", "с", "к", "со", "ко", "про", "до")

2. words substitution: will ask to submit a sample file with the words you want to substitute and erase them. File example attached.

3. words shuffle: will shuffle the words of all sentences. No addictional file required. 

4. verbs tenses: will put between brackets all verbs in their infinitive form. The file VERB_LIST.xlsx needs to be given as input, or if the user wants to substitute only some verbs he can pass his own file.

5. NSV-SV verbs all: will put between brackets all verbs in their infinitive form and if they are a nsv-sv pair will put both variants. The file VERB_LIST.xlsx needs to be given as input as well as VERB_PAIRS.xlsx, or if the user wants to substitute only some verbs he can pass his own file.

5. NSV-SV verbs only: will put between brackets verbs in their infinitive form only if they are a nsv-sv pair will put both variants. The file VERB_LIST.xlsx needs to be given as input as well as VERB_PAIRS.xlsx, or if the user wants to substitute only some verbs he can pass his own file.


# Word file requirements
The word file uploaded should be in .DOCX format and preferibily without special formatting. If the text is copied from other files and/or websites it is better to use the option PASTE TEXT ONLY and do not keep the source formatting, since some formatting may cause problems in the document parsing.


# Regarding verbs declination excel files
This file have been generated from dictionaries automatically and, even though it has been tested on numerous texts of different type for errers detection and correction, it is still possible to meet some rare spelling errors and/or some missing verbs. If you find an error, please let us know so we can remove it ASAP. Also in the file there are some non-existing verbs forms, since they were generated in automatic, but since they do not exist in Russian, they do not create any problem. Teacher/professor that want to help perfectioning the program may delete these forms from the excel file and send it back.

# Regarding verbs pairs NSV/SV excel files
This file have been generated from dictionaries automatically and, even though it has been tested on numerous texts of different type for errers detection and correction, it is still possible to meet some rare spelling errors. Some verbs pairs (NSV/SV) are probably missing, with time the file VERB_PAIRS.xlsx will be updated including more and more pairs. If you find an error or a missing pair, please let us know so we can remove/correct it ASAP. Pairs can be added to the ecxel file VERB_PAIRS.xlsx 
