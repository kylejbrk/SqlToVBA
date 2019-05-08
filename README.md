# SqlToVBA
A python script that converts SQL Queries to VBA strings, and vice versa.

# Description
This script takes a SQL statement, and converts it VBA strings so it can be pasted into a Module/Macro. It also takes VBA's line continuation limit into consideration by storing long queries into multiple variables.

# Requirements
pyperclip https://pypi.org/project/pyperclip/

# Usage
Copy the Query/String to your clipboard. Run one of the script using one of the following commands below.
```
SqlToVba.py -command

Commands:
-v    to convert SQL Query to VBA
-s    to convet VBA String to SQL Query
```
