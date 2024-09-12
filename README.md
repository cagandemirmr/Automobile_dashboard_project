# ABOUT THE DATASET
Dataset is consist of 865 rows and 8 columns.

Columns called CODE,STATE,SALERS,PRODUCTS,NET SALE AMOUNT,TARGET SALE AMOUNT,TARGET REALIZATION and YEAR.

CODE is abbreviation of STATES of USA which have 50 states.

SALERS has total 59 salesman in these States.In some states, there are more than one salesman.

PRODUCT column has 5 variables which are Mercedes,BMW,Ford,Mazda,Chevrolet.

YEARS column has three values which are 2018,2019 and 2020.

# DESIGN DECISIONS
To design dashboard, first i define borders of design.Thus, i close Guides and Highlights in Excel and select maximum areas with tolerance.

The main principle in defining borders of dashboard is doing boarding process when zoom is 100%

![image](https://github.com/user-attachments/assets/06e3273f-055d-4496-ab80-93bea8cb41d8)

In color, I personally like blue and white because it gives feel of trust and calms people and look classy as well.

Also, Because of one page constain, i would like to create pop up windows and dashboards using macros.It wil consist of three pages, in one main page to show general picture of net sales and in two pop up pages, people can observe products sales by states and performance of salesman by target values.

In the main page, i would like to show years in list box to define years to observe total sales of products KPIS. Also,On the right i would like to see bar graph to observe which product is the most sales in the selected year.Next,i want to show KPIS that people can
see total sales and Total target of the selected  year. Also i want to add graphic to show performance.And In the middle of dashboard people can see General map of USA with net sales and STATE name on it. On the bottom people can observe sales of products by years and total.

In terms of pop up pages,they can be hidden by close button and can be opened by buttons on dashboard.Pop up pages can be called as Region Report and Salesman Report and can be visible buttons on top of dashboard.
In Region Report,it shows general performance of products by sales and have two main list box which are STATES and YEAR to calculate real,target values and T/R ratio.
In Salesman Report, it has three listbox and shows total sales by products.

# PREPERATION
Before working on datas, i collect all PNG files related to brands in images sheet on Excel.
![image](https://github.com/user-attachments/assets/c9eb7d1f-aa78-4d84-b8bf-df852c92a077)

and create Parameters and Dashboard sheets on Excel.
I will use Parameters sheet for creating tables.


# PROCESS
First of all, i checked dataset for duplicates, missing values and  wrong formatted values.But dataset was clean thus i convert datas into table.
![image](https://github.com/user-attachments/assets/34f6e7a4-b0db-4255-a091-6cf48e4a05e6)

Secondly, i need unique values such as STATES, YEAR and PRODUCTS to create tables and listbox. Thus, i use UNIQUE function to define these values.
Next, i select and rename STATES as "Dönemler".And i opened Developer tab to create list box under Add tab >> Active X.

![image](https://github.com/user-attachments/assets/9b38785a-33f0-486d-a387-377f468233eb)

Under Properities section in listbox,i assign "Dönemler" to "ListFillRange" and assign output to "LinkedCell" in properities.

![image](https://github.com/user-attachments/assets/74ec11b0-a31c-4179-ade6-778535e94198)

And i create  values by SUMIFS, as a total value, i choose net sales amount,as a statement, i choose  YEARS in DATA and "Dönemler" output.Also, i choose STATES in DATA sheet and unique STATES in parameter sheet to adjust total values by YEAR and STATE values.

![image](https://github.com/user-attachments/assets/fa57cad6-0650-4d10-ab81-96ff9af077db)

Based on this information, i create kartogram.
![image](https://github.com/user-attachments/assets/dfc17b3c-aae0-43e9-a294-cad636b9067f)













