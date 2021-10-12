# -*- coding:utf-8 -*-
import sys
from excel import (
    ExcelOpt   
)
if __name__ == '__main__':
    a = """

    **************************************************************
    *                                                            *
    *    .=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-.      *
    *    |                     ______                     |      *
    *    |                  .-"      "-.                  |      *
    *    |                 /            \                 |      *
    *    |     _          |              |          _     |      *
    *    |    ( \         |,  .-.  .-.  ,|         / )    |      *
    *    |     > "=._     | )(__/  \__)( |     _.=" <     |      *
    *    |    (_/"=._"=._ |/     /\     \| _.="_.="\_)    |      *
    *    |           "=._"(_     ^^     _)"_.="           |      *
    *    |               "=\__|IIIIII|__/="               |      *
    *    |              _.="| \IIIIII/ |"=._              |      *
    *    |    _     _.="_.="\          /"=._"=._     _    |      *
    *    |   ( \_.="_.="     `--------`     "=._"=._/ )   |      *
    *    |    > _.="                            "=._ <    |      *
    *    |   (_/                                    \_)   |      *
    *    |                                                |      *
    *    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='      *
    *                                                            *
    *           LASCIATE OGNI SPERANZA, VOI CH'ENTRATE           *
    **************************************************************                                                  
                                               
功能列表：                                                                                
 1.提取出办事处、客户编码、订单量表
 2.提取费用网点明细表
    """
    print(a)
    choice_function = input('请选择:')
    if choice_function == '1':
        opt = ExcelOpt()
        opt.extractData()
    elif choice_function == '2':
        opt = ExcelOpt()
        opt.extractDataDetail()
    else:
        print('没有此功能')
        sys.exit(1)