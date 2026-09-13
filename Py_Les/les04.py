#String data type

#literal assignment
first = 'Mariam'
last = 'Sulleiman'

# print(type(first))
# print(type(first) == str)
# print(isinstance(first, str))

# constructor function
# pizza = str("pepperoni")
# print(type(pizza))
# print(type(pizza) == str)
# print(isinstance(pizza, str))

#concatenation
fullname = first + " " + last
print(fullname)

fullname += "!"
print(fullname)

#casting a number to a string
decade = str(1990)
print(type(decade))
print(decade)

statement = fullname + " " + "was born in" + " " + decade
print(statement)

multiline = """
Hey, how are you?    

I was just checking in.   

                            All good?

"""
print (multiline)

#Escaping special characters

sentence = 'I\'m back at work!\t Hey! \n\n Where\'s this at \\located?'
print(sentence)

#string methods
print(first)
print(first.lower())
print(first.upper())
print(first)

print(multiline.title)
print(multiline.replace("good", "ok"))
print(multiline)

print(len(multiline))
multiline += "                                         "
multiline = "             " + multiline
print(len(multiline))

#remove whitespace
print(len(multiline.strip()))
print(len(multiline.lstrip()))
print(len(multiline.rstrip()))

#Build a menu
title = 'menu'.upper()
print(title.center(20, '='))
print('coffee'.ljust(16, '.') + '$1'.rjust(4))
print('Muffin'.ljust(16, '.') + '$2'.rjust(4))
print('Cheesecake'.ljust(16, '.') + '$4'.rjust(4))

#string index values
print(first[0])
print(first[-1])
print(first[1:-1])
print(first[1:])

#Some methods return boolean data
print(first.startswith('M'))
print(first.endswith('Z'))

#Boolean data type
myvalue = True
x = bool(False)
print(type(x))

#Numeric Data types

price = 100
best_price = int(80)
print(type(price))
print(isinstance(best_price, int))

#float type
gpa = 3.28
y = float(1.14)
print(type(gpa))
print(type(y))

#complex type
comp_value = 5 + 3j
print(type(comp_value))
print(comp_value.real)
print(comp_value.imag)

#Built-in functions for numbers
print(abs(gpa))

print(round(gpa))

print(round(gpa, 1))

import math

print(math.pi)
print(math.sqrt(64))
print(math.ceil(gpa))
print(math.floor(gpa))

#casting string to a number
zipcode = "10001"
zip_value = int(zipcode)
print(type(zip_value))