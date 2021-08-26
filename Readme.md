安装所需库

```
pip install requests
pip install websockets==6.0
```

配置个人信息

如果还未有API，可[点击](https://www.okex.com/account/users/myApi)前往官网进行申请

```
api_key = ""
secret_key = ""
passphrase = ""
```

1.将个人信息填写到get_balance.py的init_basics函数中

2.第一次使用请首先调用init_account_order函数，获取近三个月的历史订单信息（OKex只提供近三个月），以后每次更新数据调用update函数即可

![image-20210826094424782](/okex/image-20210826094424782.png)

3.check_down函数用于查看所有亏损币种的订单信息

   check_up函数用于查看所有盈利币种的订单信息

   check_one_coin函数用于查看单一币种的订单信息

4.官方提供的API接口功能很多，可自行二次开发更多功能。

5.官方不提供实时的USDT价格，所以默认的价格为6.45![image-20210826094727270](C:\Users\xyk\Documents\GitHub\XiaoKeKeLa.github.io\posts\2020\10\18\image-20210826094727270-1629943124515.png)

6.安装过程中提示某种库不存在，可自行根据提示添加

