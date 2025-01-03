This is a word search puzzle generator. Users can customize some variables.

### How to use it?

The word_list.txt file is from https://www.mit.edu/~ecprice/wordlist.10000. If you have another word lists, you can modify the word_list.txt file. Then just run generator.py program with python.

```
python generator.py
```

### What can you control?

- Size of the puzzle
    - Check the line 76 and 118. The default size is 16x16.
- Number and length of words
    - e.g. 2 words with length 4, 3 words with length 5, ...
    - Check the line from 79 to 83.
- Number of each directions
    - e.g. 2 for S, 1 for NW, 1 for W, ...
    - Check the line 84.
- Alphabet distribution in empty space
    - The empty space is filled with random alphabets that follow the distribution of word list.
    - The function `calculate_letter_frequencies` and `fill_empty_spaces` are about it.

### Result

The puzzles are generated as docx file and word lists are saved as json file. You can also include the word list under the puzzle if you uncomment the line 135~139.

<img width="400" alt="image" src="https://github.com/user-attachments/assets/24f36cac-d89e-4397-b4bc-8420ed19094f" />

docx puzzle file

<img width="100" alt="image" src="https://github.com/user-attachments/assets/9e30a1e7-03b6-4bf9-b29b-7e54b0a5de95" />

json word list file
