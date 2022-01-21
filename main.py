# -*- coding:utf-8 -*-
import sys
from excel import (
    ExcelOpt,ExcelOptNew  
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
 1.提取出办事处、客户编码、订单量表（old）
 2.提取费用网点明细表（old）
 3.提取出办事处、客户编码、订单量表（new）
 4.提取费用网点明细表（new）
    """
    print(a)
    choice_function = input('请选择:')
    if choice_function == '1':
        opt = ExcelOpt()
        opt.extractData()
    elif choice_function == '2':
        opt = ExcelOpt()
        opt.extractDataDetail()
    elif choice_function == '3':
        opt = ExcelOptNew()
        opt.extractData()
    elif choice_function == '4':
        opt = ExcelOptNew()
        opt.extractDataDetail()
    else:
        print('没有此功能')
        sys.exit(1)