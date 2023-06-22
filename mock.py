import pandas as pd
import random
import string

# 生成随机字符串工号


def generate_employee_id():
    return ''.join(random.choices(string.digits, k=10))

# 生成随机汉字姓名


def generate_employee_name():
    n = random.randint(2, 3)
    name = ''.join(random.sample('人之初性本善性相近习相远苟不教性乃迁教之道贵以专', n))
    return name


# 创建 DataFrame
data = {'工号': [generate_employee_id() for _ in range(100)],
        '姓名': [generate_employee_name() for _ in range(100)]}
df = pd.DataFrame(data)

# 导出到 Excel
df.to_excel('employee_data.xlsx', index=False)
