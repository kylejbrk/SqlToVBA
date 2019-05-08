#! python3
# Converts SQL Queries to a multi lined VBA string, or converts a multi line VBA string back to a SQL Query

import pyperclip, sys

def main():
    convertedText = ''
    try:
        if sys.argv[1] == '-v':
            convertedText = converToVba(pyperclip.paste())
            pyperclip.copy(convertedText)
            print('Text has been converted for VBA and copied to clipboard')
        elif sys.argv[1] == '-s':
            convertedText = converToSql(pyperclip.paste())
            pyperclip.copy(convertedText)
            print('Text has been converted for SQL and copied to clipboard')
        elif sys.argv[1] == '-h':
            print('Usage: SqlToVba.py -command\nCommands:\n\t-s\tif desired output is SQL\n\t-v\tif desired output is VBA')
        else:
            raise Exception()
    except:
        print('Please type -h for help')

def converToVba(sqlString):
    try:
        lines = sqlString.split('\n')
        stringNum = 1
        newQry = False
        for idx,line in enumerate(lines):
            if idx % 24 == 0:
                lines.insert(idx, 'Dim sqlQry' + str(stringNum) + ' as String') 
                newQry = True            
            else:    
                if newQry:
                    lines[idx] = 'sqlQry' + str(stringNum) + ' = "' + line.strip() + ' " & _'   
                    lines[idx - 2] = lines[idx - 2].replace('& _','\n')
                    stringNum += 1        
                else:    
                    lines[idx] = '\t"' + line.strip() + ' " & _'       
                newQry = False    
        lines[-1] = lines[-1].replace('& _','')              

        stringNums = []
        for x in range(1, stringNum):
            stringNums.append('sqlQry' + str(x))

        qryFinal = '\nDim sqlQryFinal As String \nsqlQryFinal = ' + ' + '.join(stringNums)
        lines.append(qryFinal)

        return '\n'.join(lines)
    except ValueError:
        print("!!! Please enter a valid string !!!")
    except Exception as e:
        print(e)


def converToSql(vbaString):
    try:
        lines = vbaString.split('\n')
        for idx,line in enumerate(lines[:]):
            line = line.strip()
            lines[idx] = line.replace('& _','').replace('&','').replace('"','')
        
        return '\n'.join(lines)
    except ValueError:
        print("!!! Please enter a valid string !!!")
    except Exception as e:
        print(e)    

if __name__ == '__main__':
    main()