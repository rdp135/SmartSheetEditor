'''

Module for handling all excel related data management

'''

import numpy as np
import pandas as pd

class xlsx_handler:

    def __init__(self):
        pass

    @staticmethod
    def gen_xlsx(x, y):
        '''
        Generate a .xlsx file with size x, y of random int in range 0-100
        :param x: length along x
        :param y: length along y
        :return: pandas dataframe
        '''
        df = pd.DataFrame(np.random.randint(0, 100, size=(x, y)))
        df.to_excel("test.xlsx", index=False, header=False)
        return df

    @staticmethod
    def get_xlsx(name):
        '''
        Read a .xlsx file given its name
        :param name: name of .xlsx file
        :return: .xlsx file as numpy array
        '''
        df = pd.read_excel(name, header=None)
        df.replace(['nan', 'None'], '')
        df_np = np.array(df)
        return df_np
