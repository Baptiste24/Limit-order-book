# -*- coding: utf-8 -*-

"""Objectives : Use Python to construct a limit order book
and simulate the dynamic trading process."""

################################# LIBRARIES #################################

from random import randint
import random #to generate random order quantities for instance
from datetime import datetime #to get the time of the creation of an order
#from pandas import DataFrame
import xlsxwriter

############################## GLOBAL VARIABLES ##############################

NB_TOTAL_ORDERS = 5000

#We set the range for the upper bound limit price (for buy limit orders)
UBPL_MIN = 5
UBPL_MAX = 10

#We set the range for the lower bound limit price (for sell limit orders)
LBPL_MIN = 5
LBPL_MAX = 10

################################# FUNCTIONS #################################

#question 3
def random_order_quantity():
    """

    Returns a random quantity (int) between 100 and 1000 (included - numbers choosen by the team) with a step of 100 to avoid having an odd number

    """
    return random.randrange(100,1001,100)

#order extecuted at the best possible prices
def create_buy_market_order(ID, order_quantity):
    """
    Function that create a buy market order at the best possible price.

    Parameters
    ----------
    ID (integer) : order ID
    order_quantity (integer) : Quantity of stocks in this order.

    Returns
    -------
    list with the caracteristics of the order :
        - 0 (boolean) : BUY
        - 0 (boolean) : MARKET ORDER
        - order_quantity (integer)
        - order_price (float with 2 decimals) : random price
        - price_limit (integer) : 0 put conventionnaly to say there is no price_limit
        - creation_time_string (string) : creation time of this order.

    """
    order_price = round(random.uniform(95, 100), 2) #generate a random price between 98 and 102 (included - numbers choosen by the team) with a step of 0.01
    
    #get the time of the creation of this order
    creation_time = datetime.now()
    
    return [ID, 0, 0, order_quantity, order_price, 0, str(creation_time)]

#order extecuted at the best possible prices
def create_sell_market_order(ID, order_quantity):
    """
    Function that create a sell market order at the best possible price.

    Parameters
    ----------
    ID (integer) : order ID
    order_quantity (integer) : Quantity of stocks in this order.

    Returns
    -------
    list with the caracteristics of the order :
        - 1 (boolean) : SELL
        - 0 (boolean) : MARKET ORDER
        - order_quantity (integer)
        - order_price (float with 2 decimals) : random price
        - price_limit (integer) : 0 put conventionnaly to say there is no price_limit
        - creation_time_string (string) : creation time of this order.

    """
    order_price = round(random.uniform(100.01, 105), 2) #generate a random price between 98 and 102 (included - numbers choosen by the team) with a step of 0.01
    
    #get the time of the creation of this order
    creation_time = datetime.now()
    return [ID, 1, 0, order_quantity, order_price, 0, str(creation_time)]

def create_buy_limit_order(ID, order_quantity):
    """
    Function that create a buy limit order.

    Parameters
    ----------
    ID (integer) : order ID
    order_quantity (integer) : Quantity of stocks in this order.

    Returns
    -------
    list with the caracteristics of the order :
        - 0 (boolean) : BUY
        - 1 (boolean) : LIMIT ORDER
        - order_quantity (integer)
        - order_price (float with 2 decimals) : random price
        - upper_bound_price_limit (float)
        - creation_time_string (string) : creation time of this order.

    """
    order_price = round(random.uniform(90, 100), 2) #generate a random price between 90 and 110 (included - numbers choosen by the team) with a step of 0.01
    upper_bound_price_limit = order_price + randint(UBPL_MIN, UBPL_MAX + 1) #upper bound price limit formula for BUY ORDER, choosen by the team, between 3 and 10 added to the order price
    
    #get the time of the creation of this order
    creation_time = datetime.now()
    return [ID, 0, 1, order_quantity, order_price, upper_bound_price_limit, str(creation_time)]

def create_sell_limit_order(ID, order_quantity):
    """
    Function that create a sell limit order.

    Parameters
    ----------
    ID (integer) : order ID
    order_quantity (integer) : Quantity of stocks in this order.

    Returns
    -------
    list with the caracteristics of the order :
        - 1 (boolean) : SELL
        - 1 (boolean) : LIMIT ORDER
        - order_quantity (integer)
        - order_price (float with 2 decimals) : random price
        - lower_bound_price_limit (float)
        - creation_time_string (string) : creation time of this order.

    """
    order_price = round(random.uniform(100.01, 110), 2) #generate a random price between 90 and 110 (included - numbers choosen by the team) with a step of 0.01
    lower_bound_price_limit = order_price - randint(LBPL_MIN, LBPL_MAX + 1) #lower bound price limit formula for BUY ORDER, choosen by the team, between LBPL_MIN and LBPL_MAX substracted to the order price
    
    #get the time of the creation of this order
    creation_time = datetime.now()
    return [ID, 1, 1, order_quantity, order_price, lower_bound_price_limit, str(creation_time)]


#question 5
def takeID(elem):
    """
    Function that returns the ID of an order

    Returns
    -------
    int
      0th element (=the ID) of an order.

    """
    return elem[0]

def takeOrderDirection(elem):
    """
    Function that returns 0 if it is a buy order or 1 if it is a sell order

    Returns
    -------
    boolean
        2nd element (=the buy/sell) of an order.

    """
    return elem[1]

def takeOrderType(elem):
    """
    Function that returns 0 if it is a market order or 1 if it is a limit order

    Returns
    -------
    boolean
        3rd element (=the limit/market) of an order.

    """
    return elem[2]

def takeQuantity(elem):
    """
    Function that returns the 4th element (=the quantity) of an order.

    Returns
    -------
    int
        4th element (=the quantity) of an order.

    """
    return elem[3]

def takePrice(elem):
    """
    Function that returns the 5th element (=the price) of an order

    Returns
    -------
    float
        5th element (=the price) of an order.

    """
    return elem[4]

def takePriceLimit(elem):
    """
    Function that returns the 6th element (=the price) of an order

    Returns
    -------
    float
        6th element (=the price) of an order.

    """
    return elem[5]

def takeTime(elem):
    """
    Function that returns the 7th element (=the time) of an order

    Returns
    -------
    string
        7th element (=the time) of an order.

    """
    return elem[6]


#Question 7
def display_Top10_orders(list_buy_orders, list_sell_orders):
    """
    Displays the 10 first buy and sell orders from the order book

    Parameters
    ----------
    list_buy_orders : list
        list of all the unexecuted buy orders
    list_sell_orders : list
        list of all the unexecuted sell orders

    Returns
    -------
    None.

    """
    print("10 best buy and sell orders are : ")
    for i in range(10):
        print("buy order n°" + str(i+1) + ":" + str(list_buy_orders[i]) + "\nsell order n°" + str(i+1) + ":" + str(list_sell_orders[i]) + "\n")




#Question 8
def generate5000orders(list_buy_orders, list_sell_orders):
    """
    Function that generate 5000 buy/sell limit/market orders.

    Parameters
    ----------
    list_buy_orders : list of limit/market buy orders
        This list is set in this function.
    list_sell_orders : list of limit/market sell orders
        This list is set in this function.

    Returns
    -------
    None.

    """
    i_buy = 0
    i_sell = 0
    nb_limit_orders = randint(2000, 3000) #Value decided in order to have almost as much limit orders than market orders
    
    for i in range(nb_limit_orders):
        buy_or_sell = randint(0, 1)
        if buy_or_sell == 0: #if buy
            list_buy_orders.append(create_buy_limit_order(i_buy, random_order_quantity()))
            i_buy += 1
        else: #if sell
            list_sell_orders.append(create_sell_limit_order(i_sell, random_order_quantity()))
            i_sell += 1
    
    for i in range(nb_limit_orders, NB_TOTAL_ORDERS): #nb market_orders
        buy_or_sell = randint(0, 1)
        if buy_or_sell == 0: #if buy
            list_buy_orders.append(create_buy_market_order(i_buy, random_order_quantity()))
            i_buy += 1
        else: #if sell
            list_sell_orders.append(create_sell_market_order(i_sell, random_order_quantity()))
            i_sell += 1
    #IDs go from 0 to nb_limit_orders and from 0 to nb_market_orders


#question 11
def bid_ask_spread(bid, ask):
    """
    Function that calculate the bid ask spread between two orders given the bid price and the offer price.

    Parameters
    ----------
    bid : float
        bid price.
    ask : float
        ask/offer price.

    Returns
    -------
    positive float
        Return the bid ask spread.

    """
    spread = abs(bid - ask)
    return spread


################################ MAIN PROGRAM ################################

list_buy_orders = []
list_sell_orders = []

generate5000orders(list_buy_orders, list_sell_orders)

"""
print(list_buy_orders)
print("\n\n\n\n")
print(list_sell_orders)
print("\n\n\n\n")
"""


# sort lists buy and sell orders by market then limit order, then by price, and if two orders have the same price, sort by limit price and finally by time priority
list_buy_orders.sort(key=takeTime)
list_buy_orders.sort(key = takePrice, reverse = True)
list_buy_orders.sort(key=takePriceLimit, reverse = True)
list_buy_orders.sort(key = takeOrderType, reverse = False)

list_sell_orders.sort(key=takeTime)
list_sell_orders.sort(key = takePrice, reverse = False)
list_sell_orders.sort(key=takePriceLimit, reverse = False)
list_sell_orders.sort(key = takeOrderType, reverse = False)

# print list
#print('Sorted list:', list_buy_orders, '\n\n\n\n\n\nSorted list:', list_sell_orders)



execution_price_file = open("execution price.txt", "w")
execution_price_file.write("BUY/SELL	price order    order ID\n")

bid_ask_spread_file = open("bid-ask spread.txt", "w")
bid_ask_spread_file.write("ID Buy order\tID sell order\tbid-ask spead\n")


bid_ask_spread_list = [] #creating this variable for Excel

i = 0
j = 0
while i < len(list_buy_orders):
    while j < len(list_sell_orders):
        
        #Q12 : compute the bid-ask spread and output the computed bid-ask spead to an external text file
        spread = bid_ask_spread(takePrice(list_buy_orders[i]), takePrice(list_sell_orders[j]))
        bid_ask_spread_file.write(str(takeID(list_buy_orders[i])) + "\t\t\t" + str(takeID(list_sell_orders[j])) + "\t\t\t" + str(spread) + "\n")
        bid_ask_spread_list.append(spread)
        
        #limit limit
        if takeOrderType(list_buy_orders[i]) == 1 and takeOrderType(list_sell_orders[j]) == 1:
            if (takePriceLimit(list_buy_orders[i]) > takePrice(list_sell_orders[j])) and (takePriceLimit(list_sell_orders[j]) < takePrice(list_buy_orders[i])):
                #trade
                #calculate the minimum quantity between the ith buy order quantity and the jth sell order quantity
                min_quantity = min(takeQuantity(list_buy_orders[i]), takeQuantity(list_sell_orders[j]))
                
                #trade execution
                list_buy_orders[i][3] -= min_quantity
                list_sell_orders[j][3] -= min_quantity
                
                if list_buy_orders[i][3] == 0:
                    execution_price_file.write("BUY\t\t" + str(takePrice(list_buy_orders[i])) + "\t\t\t" + str(takeID(list_buy_orders[i])) + "\n")
                    list_buy_orders.pop(i)
                    
                    if list_sell_orders[j][3] == 0:
                        execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                        list_sell_orders.pop(j)
                      
                else:
                    execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                    list_sell_orders.pop(j)
                        
            else:
                #no trade
                #look a the next buy and sell orders
                i += 1
                j += 1
        
        #limit market
        elif takeOrderType(list_buy_orders[i]) == 1 and takeOrderType(list_sell_orders[i]) == 0:
            if takePriceLimit(list_buy_orders[i]) > takePrice(list_sell_orders[j]):
                #trade
                #calculate the minimum quantity between the ith buy order quantity and the jth sell order quantity
                min_quantity = min(takeQuantity(list_buy_orders[i]), takeQuantity(list_sell_orders[j]))
                
                #trade execution
                list_buy_orders[i][3] -= min_quantity
                list_sell_orders[j][3] -= min_quantity
                
                if list_buy_orders[i][3] == 0:
                    execution_price_file.write("BUY\t\t" + str(takePrice(list_buy_orders[i])) + "\t\t\t" + str(takeID(list_buy_orders[i])) + "\n")
                    list_buy_orders.pop(i)
                    
                    if list_sell_orders[j][3] == 0:
                        execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                        list_sell_orders.pop(j)
                else:
                    execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                    list_sell_orders.pop(j)
                        
            else:
                #no trade
                #look a the next buy and sell orders
                i += 1
                j += 1
        
        #market limit
        elif takeOrderType(list_buy_orders[i]) == 0 and takeOrderType(list_sell_orders[i]) == 1:
            if takePriceLimit(list_sell_orders[j]) < takePrice(list_buy_orders[i]):
                #trade
                #calculate the minimum quantity between the ith buy order quantity and the jth sell order quantity
                min_quantity = min(takeQuantity(list_buy_orders[i]), takeQuantity(list_sell_orders[j]))
                
                #trade execution
                list_buy_orders[i][3] -= min_quantity
                list_sell_orders[j][3] -= min_quantity
                
                if list_buy_orders[i][3] == 0:
                    execution_price_file.write("BUY\t\t" + str(takePrice(list_buy_orders[i])) + "\t\t\t" + str(takeID(list_buy_orders[i])) + "\n")
                    list_buy_orders.pop(i)
                    
                    if list_sell_orders[j][3] == 0:
                        execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                        list_sell_orders.pop(j)
                else:
                    execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                    list_sell_orders.pop(j)
                        
            else:
                #no trade
                #look a the next buy and sell orders
                i += 1
                j += 1
        
        #market market
        elif takeOrderType(list_buy_orders[i]) == 0 and takeOrderType(list_sell_orders[i]) == 0:
            #trade (a buy market order and a sell market order inevitably means trade)
            #calculate the minimum quantity between the ith buy order quantity and the jth sell order quantity
            min_quantity = min(takeQuantity(list_buy_orders[i]), takeQuantity(list_sell_orders[j]))
            
            #trade execution
            list_buy_orders[i][3] -= min_quantity
            list_sell_orders[j][3] -= min_quantity
            
            if list_buy_orders[i][3] == 0:
                execution_price_file.write("BUY\t\t" + str(takePrice(list_buy_orders[i])) + "\t\t\t" + str(takeID(list_buy_orders[i])) + "\n")
                list_buy_orders.pop(i)
                
                if list_sell_orders[j][3] == 0:
                    execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                    list_sell_orders.pop(j)
            else:
                execution_price_file.write("SELL\t\t" + str(takePrice(list_sell_orders[j])) + "\t\t" + str(takeID(list_sell_orders[j])) + "\n")
                list_sell_orders.pop(j)
        
        #end of the lists, no more trade possible
        if j >= len(list_sell_orders) or i >= len(list_buy_orders) or j == 0 or i == 0 :
            break
    if j >= len(list_sell_orders):
        break
        
bid_ask_spread_file.close()
execution_price_file.close()

#Order book
book_order = {"list_buy_orders" : list_buy_orders, "list_sell_orders" : list_sell_orders}

display_Top10_orders(list_buy_orders, list_sell_orders)


"""
counter_buy = 0
for i in range(len(list_buy_orders)):
    if takeOrderType(list_buy_orders[i]) == 1:
        counter_buy +=1
print(counter_buy)

counter_sell = 0
for i in range(len(list_sell_orders)):
    if takeOrderType(list_sell_orders[i]) == 1:
        counter_sell +=1
print(counter_sell)
"""

"""
#plot the graph in python plots window
df = DataFrame({'bid-ask spread': bid_ask_spread_list})
df.to_excel('python_projet.xlsx', sheet_name='bid-ask spread', index=False)
"""

#Q10 and Q13 : Plot price process and bid-ask spread in Excel

# Buy order price data location inside excel
buy_price_list = []
for i in range(len(list_buy_orders)):
    buy_price_list.append(takePrice(list_buy_orders[i]))

buy_price_data = [buy_price_list for _ in range(len(list_buy_orders))]
data_start_loc_buy = [0, 0]
data_end_loc_buy = [data_start_loc_buy[0] + len(buy_price_data), 0]

# Sell order price data location inside excel
sell_price_list = []
for i in range(len(list_sell_orders)):
    sell_price_list.append(takePrice(list_sell_orders[i]))

sell_price_data = [takePrice(list_sell_orders) for _ in range(len(list_sell_orders))]
data_start_loc_sell = [0, 1]
data_end_loc_sell = [data_start_loc_sell[0] + len(sell_price_data), 1]

# Bid ask spread data location inside excel
bid_ask_spread_data = [bid_ask_spread_list for _ in range(len(bid_ask_spread_list))]
data_start_loc = [0, 0]
data_end_loc1 = [data_start_loc[0] + len(bid_ask_spread_data), 0]


#Open Excel workbook
workbook = xlsxwriter.Workbook('python_projet.xlsx')


# Define buy orders price chart components
chart_buy_price = workbook.add_chart({'type': 'line'})
chart_buy_price.set_y_axis({'name': 'Buy orders price'})
chart_buy_price.set_x_axis({'name': 'Submitted order number'})
chart_buy_price.set_title({'name': 'Buy orders prices'})


# Define sell orders price chart components
chart_sell_price = workbook.add_chart({'type': 'line'})
chart_sell_price.set_y_axis({'name': 'Sell orders price'})
chart_sell_price.set_x_axis({'name': 'Submitted order number'})
chart_sell_price.set_title({'name': 'Sell orders prices'})


chart_sellbuy_price = workbook.add_chart({'type': 'line'})
chart_sellbuy_price.set_y_axis({'name': 'Sell orders price'})
chart_sellbuy_price.set_x_axis({'name': 'Submitted order number'})
chart_sellbuy_price.set_title({'name': 'Sell and buy orders prices'})


### worksheet setting
price_worksheet = workbook.add_worksheet("Price process") #create new worksheet
price_worksheet.write_column(*data_start_loc_buy, buy_price_list) #Adding data in Excel to plot the chart
price_worksheet.write_column(*data_start_loc_sell, sell_price_list) #Adding data in Excel to plot the chart


chart_buy_price.add_series({
    'values': [price_worksheet.name] + data_start_loc_buy + data_end_loc_buy,
    'name': "price buy orders data",
})
price_worksheet.insert_chart('D1', chart_buy_price) #insert the spread chart in cell D1


chart_sell_price.add_series({
    'values': [price_worksheet.name] + data_start_loc_sell + data_end_loc_sell,
    'name': "price sell orders data",
})

chart_sellbuy_price.add_series({
    'values': [price_worksheet.name] + data_start_loc_buy + data_end_loc_buy,
    'name': "price buy orders data",
})
chart_sellbuy_price.add_series({
    'values': [price_worksheet.name] + data_start_loc_sell + data_end_loc_sell,
    'name': "price sell orders data",
})


price_worksheet.insert_chart('L1', chart_sell_price)
price_worksheet.insert_chart('D20', chart_sellbuy_price)




# Define bid ask spread chart components
spread_chart = workbook.add_chart({'type': 'line'})
spread_chart.set_y_axis({'name': 'bid-ask spread'})
spread_chart.set_x_axis({'name': 'submitted order number'})
spread_chart.set_title({'name': 'Bid-ask spread'})

chart_sellbuy_price.set_y_axis({'min': 80, 'max': 120})
chart_sellbuy_price.set_size({'width': 900, 'height': 576})



spread_worksheet = workbook.add_worksheet("bid-ask spread") #create new worksheet
spread_worksheet.write_column(*data_start_loc, bid_ask_spread_list) #Adding data in Excel to plot the chart
spread_chart.add_series({
    'values': [spread_worksheet.name] + data_start_loc + data_end_loc1,
    'name': "bid-ask spread data",
})
spread_worksheet.insert_chart('B1', spread_chart) #insert the spread chart in cell B1

workbook.close()  # Write to file