import random
HANGMANPICS = ['''

  +---+
  |   |
      |
      |
      |
      |
=========''', '''

  +---+
  |   |
  O   |
      |
      |
      |
=========''', '''

  +---+
  |   |
  O   |
  |   |
      |
      |
=========''', '''

  +---+
  |   |
  O   |
 /|   |
      |
      |
=========''', '''

  +---+
  |   |
  O   |
 /|\  |
      |
      |
=========''', '''

  +---+
  |   |
  O   |
 /|\  |
 /    |
      |
=========''', '''

  +---+
  |   |
  O   |
 /|\  |
 / \  |
      |
=========''']

def displayBoard(HANGMANPICS, missedLetters, correctLetters, secretWord):
    print(HANGMANPICS[len(missedLetters)])
    print()

    print('Missed letters:', end=' ') #end 문을 쓰면 줄바꿈 생략
    for letter in missedLetters:
        print(letter, end=' ') #end 문을 쓰면 줄바꿈 생략
    print()

    blanks = '_' * len(secretWord)

    for i in range(len(secretWord)): # replace blanks with correctly guessed letters
        if secretWord[i] in correctLetters:
            blanks = blanks[:i] + secretWord[i] + blanks[i+1:]

    for letter in blanks: # show the secret word with spaces in between each letter
        print(letter, end=' ') #end 문을 쓰면 줄바꿈 생략
    print()

words = 'ant baboon badger bat bear beaver camel cat clam cobra cougar coyote crow deer dog donkey duck eagle ferret fox frog goat goose hawk lion lizard llama mole monkey moose mouse mule newt otter owl panda parrot pigeon python rabbit ram rat raven rhino salmon seal shark sheep skunk sloth snake spider stork swan tiger toad trout turkey turtle weasel whale wolf wombat zebra'.split()

def getRandomWord(wordList):
# This function returns a random string from the passed list of strings.
    wordIndex = random.randint(0, len(wordList) - 1)
    return wordList[wordIndex]

missedLetters = ''
correctLetters = ''
secretWord = getRandomWord(words)

print('Missed letters:',)
for letter in missedLetters:
    print(letter)

blanks = '_' * len(secretWord)

print ('secretWord:',secretWord)
print ('blanks:', blanks)

for i in range(len(secretWord)):  # replace blanks with correctly guessed letters
    if secretWord[i] in correctLetters:
        blanks = blanks[:i] + secretWord[i] + blanks[i + 1:]

print (blanks)
print (len(secretWord))
print (correctLetters)

for letter in blanks:  # show the secret word with spaces in between each letter
    print(letter, end=' ')  # end 문을 쓰면 줄바꿈 생략

