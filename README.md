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


I repeat same process with SUMIF function for the total sales of PRODUCT's KPI according to selected YEAR.

![image](https://github.com/user-attachments/assets/44786c77-5f69-45bb-9075-699f3e13b0a3)


Thus i assign each value related to PRODUCT.
![image](https://github.com/user-attachments/assets/3e6d5eec-d1ab-4480-89cc-3b9667b03928)

Also, create bar graph from this table.

![image](https://github.com/user-attachments/assets/461553b9-c91b-41eb-90d8-9b2de159cebf)

I repeat same process for the total sales and total target sales by YEAR.So i create table to show their KPI in the Dashboard.
![image](https://github.com/user-attachments/assets/2ba48e0a-5ba8-465e-aac0-4f86002a2b6b)

![image](https://github.com/user-attachments/assets/d7898a9f-498a-4819-b6b2-d89bb9021aa5)

In addition to that i want to create indicator.Thus i need to define rotation limit for my indicator.First i divide Total sales by Total target in F4 cell.Then i write function "IF(F14>=1;1;F14)" that if value is beyond 100 per cent it will stop at 1 otherwise it will show value in the cell.
![image](https://github.com/user-attachments/assets/59d80528-9935-47fa-b543-5c8d5b2f3ef8)

First I create arrow then allign to the middle of indicator panel.And rename it as "ibre".
In the developer panel in Excel, i activate the creaor mode and assign makro to the "ibre" and write the code below.With these my indicator panel works will change by year.
![image](https://github.com/user-attachments/assets/3e9751bb-b464-42ce-ac56-469f56690864)

In the last part in the main dashboard, i create panel for stable graphic to show changes by using unique values in PRODUCTS and YEAR.
![image](https://github.com/user-attachments/assets/855dbf58-605d-48ae-853f-d21b59861574)

![image](https://github.com/user-attachments/assets/05787c89-6bd1-4dc1-9044-38bcf7c9d74f)

At the and of dessigning dashboard, i add buttons for "Salesman Report" and "Region Report".
It looks like this in the end of Dashboard design.
![image](https://github.com/user-attachments/assets/f1725bd3-c5cb-4bf5-95cf-63175f43bc4b)

All process was the same in exept closing and opening windows.In "Salesman Report" , i lastly design closing cross, define all dashboard as "bolge" and define macro and write code like below.

![image](https://github.com/user-attachments/assets/1cec8652-d0c4-444c-a19f-be9e0332c2f7)

I define same code to the button with little changes like using "True" statement instead af "False" statement.

![image](https://github.com/user-attachments/assets/79747a50-4dc2-4d79-af5e-6ae6c8b2ab9a)

![image](https://github.com/user-attachments/assets/5747a22d-6d65-4caa-9e57-c8bde3952c15)


I reapeat same process in "Region Report".
![image](https://github.com/user-attachments/assets/7afa86b3-a1c6-4b8d-9147-894cb6112d65)





















