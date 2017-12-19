# cryptocurrencies2excel

This program allows you to generate a spreadsheet from updated crypto-currencies information with live data.
The exported spreadsheet is generated from a given template.
This allows you to configure your template to calculate compound data automatically, to show data in a custom graph, etc. 

### Source ###

The program is based on [coinmarketcap.com](https://coinmarketcap.com) API.

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

### Configuration ###

Configuration is only applied to `Wallets` sheet data.
In order to configure the tool, you have to edit the [wallets.json](https://github.com/auino/cryptocurrencies2excel/blob/master/wallets.json) file by including data about your wallets.
Data are expressed in JSON format and you have to specify, for each wallet, the symbol (case-sensive; as generated on the `Data` sheet) and the amount of coins you own.

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
