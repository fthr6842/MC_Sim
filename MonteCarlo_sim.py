import numpy as np
import pandas as pd
import sys
#sys.argv
# 參數設定
print(sys.argv)
code = str(sys.argv[1])
S0 = float(sys.argv[2]) # 初始股價
mu = float(sys.argv[3]) # 年化收益率
sigma = float(sys.argv[4]) # 年化波動率
T = int(sys.argv[5]) # 模擬期限（天）
M = int(sys.argv[6]) # 模擬路徑數量
Td = int(sys.argv[7]) #全年交易天期      
T = T/M
dt = 1/M # 時間步長(年)
N = int(T / dt) # 時間步數
if len(sys.argv) == 8:
    Dir = ""
else:
    Dir = str(sys.argv[8])     

def simulate_stock_price(S0, mu, sigma, T, dt, N):
    """模擬股價路徑"""
    S = np.zeros(N)
    S[0] = S0
    for t in range(1, N):
        Z = np.random.normal()  # 標準正態分佈隨機數
        S[t] = S[t-1] * np.exp((mu - 0.5 * sigma**2) * dt + sigma * np.sqrt(dt) * Z)
    return S

df = pd.DataFrame()

for i in range(M):
    stock_path = simulate_stock_price(S0, mu, sigma, T, dt, N)
    df_temp = pd.DataFrame(stock_path)
    df = pd.concat([df, df_temp], axis = 1)

df.columns = [str(i+1) for i in range(M)]
if Dir == "":
    df.to_excel('output_' + code + '.xlsx')
else:
    df.to_excel(Dir + '/output_' + code + '.xlsx')
