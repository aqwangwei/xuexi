# 1. 导入工具
import pandas as pd

# 2. 准备数据（这就是模拟爬下来的邮件信息）
# 这里我们造 3 条假数据
mail_data = [
    ["张三", "zhangsan@test.com", "你好，吃饭了吗？"],
    ["李四", "lisi@test.com", "记得明天来开会。"],
    ["王五", "wangwu@test.com", "项目已完成。"]
]

# 3. 装进 Excel 表格
# 定义表格头
df = pd.DataFrame(mail_data, columns=["发件人", "邮箱地址", "邮件内容"])

# 4. 保存文件
# 这个文件会保存在电脑的“文档”文件夹里
df.to_excel("C:/Users/Administrator/Documents/邮件数据.xlsx", index=False)

print("恭喜！Excel 文件生成成功了！快去文档文件夹看看。")
