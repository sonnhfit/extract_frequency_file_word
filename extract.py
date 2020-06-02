import docx2txt
import pandas as pd

#file word định dạng docx
my_text = docx2txt.process("OECD-converted.docx")

# Python code to find frequency of each word 
def freq(str):
    str = str.lower()
  
    # break the string into list of words  
    str = str.split()          
    str2 = [] 
  
    # loop till string values present in list str 
    for i in str:              
  
        # checking for the duplicacy 
        if i not in str2: 
  
            # insert value in str2 
            str2.append(i)  

    arr = []   
    for i in range(0, len(str2)):
        sub_ar = [str2[i], str.count(str2[i])]
        arr.append(sub_ar)
    
    df = pd.DataFrame(arr)
    df.to_csv('sondeptroai.csv', index=False, header=False)    
  
def main(): 
    str = my_text
    freq(str)                     
  
if __name__=="__main__": 
    main()             # call main function 