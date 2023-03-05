# 2.1 Python解释器


```python
print('hello,world')
```

    hello,world
    


```python
import numpy as np
a = np.random.randn(100,100)
%timeit np.dot(a,a)
```

    54.8 µs ± 1.95 µs per loop (mean ± std. dev. of 7 runs, 10000 loops each)
    


```python
%timeit np.dot(a,a)
```

    52 µs ± 1.51 µs per loop (mean ± std. dev. of 7 runs, 10000 loops each)
    


```python
%pwd
```




    'C:\\Users\\ninel\\OneDrive\\桌面\\Jupyter'




```python
%quickref
```


```python
%run 蟒蛇绘制.py
```


```python
import matplotlib.pyplot as plt
plt.plot(np.random.randn(50).cumsum())
```




    [<matplotlib.lines.Line2D at 0x1b5a5df5a08>]




![png](output_7_1.png)



```python
import numpy as np
data  = {i:np.random.randn() for i in range(7)}
```

## 2.2 IPthon基础


```python
data
```




    {0: 0.25081468401043383,
     1: 1.5260597381102907,
     2: -0.08967569133082505,
     3: -0.8022379745558693,
     4: 0.4874188307035305,
     5: 1.751528543389823,
     6: 0.5483659268869635}




```python
np.random??
```


```python
%pwd
```




    'C:\\Users\\ninel\\OneDrive\\桌面\\Jupyter'




```python
%pwd
```




    'C:\\Users\\ninel\\OneDrive\\桌面\\Jupyter'




```python
x = list(range(10))
```


```python
x
```




    [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]




```python
w = ((5,3), (4,4),(5,))
```


```python
print(w)
```

    ((5, 3), (4, 4), (5,))
    


```python
w.count(5)
```




    0




```python
def function(name, height):
    print("my" +str(height) +name)
```


```python
function(name='wahh', height=185)
```

    my185wahh
    


```python
w?
```


```python
function?
```


```python
function??
```


```python
np.random?
```

    Object `np.random` not found.
    


```python
%debug?
```


```python
%quickref
```


```python
%matplotlib inline

```


```python
import matplotlib.pyplot as plt
import numpy as np
plt.plot(np.random.randn(50).cumsum())
```




    [<matplotlib.lines.Line2D at 0x1d44c1a4fc8>]




![png](output_28_1.png)



```python
for x in w:
    if len(x) <= 1:
        print(x)
    else:
        print('setorry')
```

    setorry
    setorry
    (5,)
    


```python
type(w)
```




    tuple




```python
a = 'ninelucas'
# 字符串和元组是不可变对象，但是当不可变对象中包含可变元素的时候，它们也是可变的
a_list = ['fool', 14, [3,8]]
a_list[2]=[9,9]
a_list
```




    ['fool', 14, [9, 9]]




```python
None?
```


```python
from datetime import datetime
datetime.now()
```




    datetime.datetime(2020, 7, 3, 8, 27, 30, 5679)




```python
now = datetime.now()
```


```python
now
```




    datetime.datetime(2020, 7, 3, 8, 28, 35, 39460)




```python
type(now)
```




    datetime.datetime




```python
today = datetime.utcnow()
today
```




    datetime.datetime(2020, 7, 3, 0, 30, 15, 807845)




```python
today
```




    datetime.datetime(2020, 7, 3, 0, 30, 15, 807845)




```python
dt = datetime(2020, 8, 29)
```


```python
dt
```




    datetime.datetime(2020, 8, 29, 0, 0)




```python
someday = datetime(1997, 11, 5, 5,19,33)
```


```python
someday
```




    datetime.datetime(1997, 11, 5, 5, 19)




```python
someday.date()
```




    datetime.date(1997, 11, 5)




```python
someday.time()
```




    datetime.time(5, 19)




```python
someday.month
```




    11




```python
someday.year; someday.day
```




    5




```python
someday.strftime('%Y-%m-%d %H:%M:%S')
# 格式化datetime对象的时候，左侧自动补零
```




    '1997-11-05 05:19:33'




```python
print('今天是{0:%m}月{0:%d}日，天气不错'.format(now))
```

    今天是07月03日，天气不错
    


```python
datetime.strptime(someday.strftime('%Y-%m-%d %H:%M:%S'),'%Y-%m-%d %H:%M:%S')
```




    datetime.datetime(1997, 11, 5, 5, 19, 33)




```python
now.strftime('%F %D')
```




    '2020-07-03 07/03/20'




```python
delta = now - someday
```


```python
delta
```




    datetime.timedelta(days=8276, seconds=11342, microseconds=39460)




```python
# strftime()英语在输入上，将一个datetime.datetime类转化成字符串输出
```


```python
"""
strftime()应用在输出上，将一个datetime.datetime类转化成字符串输出
strptime()应用在输入上，当计算机读取.csv格式的数据，日期是字符串形式，可以通过strptime()转换成函数可以处理的datetime.datetime形式
strptime()是一个类方法  datetime.strptime('2020-4-13','%Y-%m-%d')
"""
```




    "\nstrftime()应用在输出上，将一个datetime.datetime类转化成字符串输出\nstrptime()应用在输入上，当计算机读取.csv格式的数据，日期是字符串形式，可以通过strptime()转换成函数可以处理的datetime.datetime形式\nstrptime()是一个类方法  datetime.strptime('2020-4-13','%Y-%m-%d')\n"




```python
datetime.strptime('2020-4-13','%Y-%m-%d')
```




    datetime.datetime(2020, 4, 13, 0, 0)




```python
"""
Bytes and characters 字节和字符
chr()将Unicode码转化成字符，ord()将字符转化为Unicode码？

"""
```




    '\nBytes and characters 字节和字符\nchr()将Unicode码转化成字符，ord()将字符转化为Unicode码？\n\n'




```python
"""链式比较"""
if ord('a') < ord('l') < ord('q'):
    print('Perfect')
```

    Perfect
    


```python
# break关键字只结束最内侧的for循环，外层的for循环会继续运行
for i in range(4):
    for j in range(4):
        if j > i :
            break
        print((j,i))
```

    (0, 0)
    (0, 1)
    (1, 1)
    (0, 2)
    (1, 2)
    (2, 2)
    (0, 3)
    (1, 3)
    (2, 3)
    (3, 3)
    


```python
# 三元表达式
result = print('Correct') if ord('r') < ord('?') else print('Incorrent')

```

    Incorrent
    


```python
v ='non-negative' if ord('r') < ord('?') else 'Negative'
v
```




    'Negative'




```python
# 第二章结束
```


```python
%pwd
```




    'C:\\Users\\ninel\\OneDrive\\桌面\\Jupyter'




```python
%pwd
```




    'C:\\Users\\ninel\\OneDrive\\桌面\\Jupyter'




```python

```
