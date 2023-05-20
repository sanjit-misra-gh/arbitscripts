import pandas as pd
import numpy as np
import time

cols = [1]
start_time = time.time()
df1 = pd.read_excel('800k.xlsx', usecols=cols)
end_time = time.time()
print(end_time - start_time)
