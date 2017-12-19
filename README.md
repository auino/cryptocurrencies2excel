# cryptocurrencies2excel

This program allows you to generate a spreadsheet from updated crypto-currencies information with live data.
The exported spreadsheet is generated from a given template.
This allows you to configure your template to calculate compound data automatically, to show data in a custom graph, etc. 

### Source ###

The program is based on [coinmarketcap.com API](https://coinmarketcap.com).

### Installation ###

 1. Clone the repository:

    ```
    git clone https://github.com/auino/cryptocurrencies2excel.git
    ```

 1. Enter the program directory:

    ```
    cd cryptocurrencies2excel
    ```

 3. Install the program requirements:

    ```
    sudo pip install -r requirements.txt
    ```

### Usage ###

In order to use the program, enter the program directory and run the following command.

```
python cryptocurrencies2excel.py
```

A new `ouput.xlsx` file will be created on the working directory, including updated data.

### External contributions ###

You are welcome to contribute this program (mainly, on the [template.xlsx](https://github.com/auino/cryptocurrencies2excel/blob/master/template.xlsx) file) in order to improve the output file format to be generated.

### Contacts ###

You can find me on Twitter as [@auino](https://twitter.com/auino).
