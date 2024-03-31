import openpyxl as xl
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

# Scap Stock Data for Fundamental Analysis Website Named Marketsmith
def Scrap(Stock):
    url = 'https://marketsmithindia.com/'

    X_path = "//*[@id='app']/div/div[2]/div[1]/div/div/div[2]/div/div/div/div/div/input"
    
    X_path_2 = "//*[@id='app']/div/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[2]/ul/li[1]/div/div[2]"
   
    X_path_3 = "//*[@id='details_placeholder_quarterlyearnings']/tr[1]"
    
    X_path_4 = "//*[@id='details_placeholder_eps']/div"
    
    X_path_5 = "//*[@id='chartTable']/tbody"
    
    X_path_6 = "//*[@id='redFlags_placeholder']/div/div[2]/table/tbody"
    
    X_path_7 = "//*[@id='company_header_placeholder']/span/div[1]/h1/b[1]"
    
    chorme = webdriver.Chrome()
    request = chorme.get(url)   
    time.sleep(2)
    Input = chorme.find_element(By.XPATH, X_path)
    M = Input.send_keys(f'{Stock}')
    time.sleep(5)
    Input.click()
    Enter = chorme.find_element(By.XPATH, X_path_2)
    V = Enter.click()
    time.sleep(10)
    L = chorme.find_element(By.XPATH, X_path_3).text
    G = L.split(' ')
    
    # Convert List Data into statment to store in excel
    Eps = G[0] + " " + "|" + G[1] + " " + "|" +  G[2] + " " + "|" +  G[3] + " " + "|" +  G[4]
    P = chorme.find_element(By.XPATH, X_path_4).text
    D = P.split('\n')
    
    # Convert List Data into statment to store in excel
    cash = D[0] + " " + "|" + D[1] + " " + "|" +  D[2] + " " + "|" +  D[3] + " " + "|" +  D[4] + " " + "|"  +  D[5]
    
    Q = chorme.find_element(By.XPATH, X_path_5).text.split('\n')
    
    # Convert List Data into statment to store in excel
    Bar = Q[0] + " " + "|" + Q[1]
    
    R = chorme.find_elements(By.XPATH, X_path_6 )
    for text in R:
        N = text.find_element(By.XPATH, "//*[@id='redFlags_placeholder']/div").text.split('\n')
    
    K = chorme.find_element(By.XPATH, X_path_7).text
    
    # Scrap Data And store in Data Disnary 
    Data = {
        "EPS": Eps,
        "Cash" : cash,
        "Barchart": Bar,
        "Red Flag" : N,
        "Com_Name" : K
        }
    return Data


# Download link of Doji scanner of stocks excel sheet Download and use in the program (Delete Limited and Ltd from the Colum B)
# https://chartink.com/screener/doji-2
Ly = xl.load_workbook("C:/Users/akaeh/Downloads/Doji, Technical Analysis Scanner (6).xlsx")
Data_1 = Ly.active

n = 3
while True:
    # Potencial Stock Name from colum B
    H = Data_1[f"B{n}"].value
    G = Scrap(H)
    Fundamentle_data = [G["EPS"], G["Cash"], G["Barchart"], G["Red Flag"], G["Com_Name"]]
    try:
        Data_1[f"H{n}"].value = Fundamentle_data[0]
    except:
        ValueError
    
    try:
        Data_1[f"I{n}"].value = Fundamentle_data[1]
        
    except:
        ValueError
        
    try:
        Data_1[f"J{n}"].value = Fundamentle_data[2]
    except:
        ValueError
        
    try:
        Data_1[f"K{n}"].value = Fundamentle_data[3][0]
    except:
        ValueError
        
    try:
        Data_1[f"L{n}"].value = Fundamentle_data[4]
    except:
        ValueError
            
    print(G)
    
    print(Data_1[f"H{n}"].value, Data_1[f"I{n}"].value, Data_1[f"J{n}"].value, Data_1[f"K{n}"].value, Data_1[f"L{n}"].value)
    
    print(Fundamentle_data)
    n += 1
    Ly.save("C:/Users/akaeh/Downloads/Doji, Technical Analysis Scanner (6).xlsx")



    

    
    
    
    
     
    


    

    
    
    
    