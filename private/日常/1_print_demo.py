# import math
# print("Dad!!")
# print("你好"+" 这是一句代码")
# print('Let\'s go')
# print("我是第一行\n我是第二行")
# print('''
# 1
# 2
# ''')
# # greet="您好，吃了么"
# # print(greet+"1")
# # print(greet+"2")
# # print(greet+"3") 全部选中，ctrl+/
#
# a = -1
# b = -2
# c = 3
# print((-b + math.sqrt(b ** 2 - 4 * a * c)) / (2 * a))
#
# s = "Hello world!"
# print(len(s))
# print(s[0])

# user_weight = float(input("请您输入您的体重（单位：kg）："))
# user_height = float(input("请您输入您的身高（单位：m）："))
# user_BMI = user_weight / (user_height) ** 2
# print(str(user_BMI))

# mood_index = int(input("心情指数："))
# if mood_index >= 60:
#     print("没问题")
# else:
#     print("没戏")

# 逻辑运算符 and or not

#列表
# shopping_list=[]
# shopping_list.append("键盘")
# shopping_list.append("鼠标")
# shopping_list.remove("键盘")
# shopping_list.append("音响")
# print(shopping_list)
# print(len(shopping_list))
# print(shopping_list[0])

# price = [799 , 1024 , 200 , 800]
# max_price = max(price)
# print(max_price)
# sorted_price = sorted(price)
# print(sorted_price)

#元组+键值对
# slang_dict = {"1":"A",
#               "2":"B"}
#
# query = (input("请查询"))
# if query in slang_dict:
#     print(slang_dict[query])
# else:
#     print("不在")

# #循环(求平均值）
# total = 0
# count = 0
# user_input = input("请输入数字，q终止程序")
# while user_input != "q":
#     num = float(user_input)
#     total += num
#     count += 1
#     user_input = input("请输入数字，q终止程序")
# result = total / count
# print("平均值为" + str(result))

#函数
# def calculate_BMI(weight, height):
#     BMI = weight / height ** 2
#     if BMI <= 18.5:
#         category = "偏瘦"
#     elif BMI <= 25:
#         category = "正常"
#     elif BMI <= 30:
#         category = "偏胖"
#     else:
#         category = "肥胖"
#     print(f"您的BMI分类为：{category}")
#     return BMI
#
# calculate_BMI(1.8, 70)

# 类
# class CuteCat:
#     def __init__(self, cat_name, cat_age, cat_color):
#         self.name = cat_name
#         self.age = cat_age
#         self.color = cat_color
# cat1 = CuteCat("Jojo",2,"yellow")

# class Student:
#     def __init__(self, name, student_id):
#         self.name = name
#         self.student_id = student_id
#         self.grades = {"语文": 0, "数学": 0, "英语": 0}
#     def set_grade(self, course, grade):
#         if course in self.grades:
#             self.grades[course] = grade
#
#     def print_grades(self):
#         print(f"学生{self.name}(学号:{self.student_id})的成绩为：")
#         for course in self.grades:
#             print(f"{course}:{self.grades[course]}分")
#
# chen = Student("小陈", "100611")
# chen.set_grade("语文", 92)
# chen.set_grade("数学", 95)
# chen.print_grades()
# zeng = Student("小曾", "100622")
# print(chen.name)
# zeng.set_grade("数学", 95)
# print(zeng.grades)

# class Employee:
#     def __init__(self, name, id):
#         self.name = name
#         self.id = id
#
#     def print_info(self):
#         print(f"员工名字：{self.name},工号：{self.id}")
#
# class FulltimeEmloyee(Employee):
#     def __init__(self, name, id, monthly_salary):
#         super().__init__(name, id)
#         self.monthly_salary = monthly_salary
#
#     def calculate_monthly_pay(self):
#         return self.monthly_salary
#
# class PartTimeEmployee(Employee):
#     def __init__(self, name, id, daily_salary, work_days):
#         super().__init__(name, id)
#         self.daily_salary = daily_salary
#         self.work_days = work_days
#
#     def calculate_monthly_pay(self):
#         return self.daily_salary * self.work_days
#
# zhangsan = FulltimeEmloyee("张三", "1001", 6000)
# lisi = PartTimeEmployee("李四", "1002", 230, 15)
# zhangsan.print_info()
# lisi.print_info()
#
# print(zhangsan.calculate_monthly_pay())
# print(lisi.calculate_monthly_pay())

# f = open("./data.txt", "r",encoding="utf-8")
# content = f.read()
# print(content)
# f.close()

# with open("./data.txt", "r",encoding="utf-8") as f:
#     print(f.readline())
#     print(f.readlines())

# with open("./poem.txt", "w", encoding="utf-8") as f:
#     f.write("我欲乘风归去，\n")
#     f.write("又恐琼楼玉宇，\n")
#     f.write("高处不胜寒。\n")

# with open("./poem.txt", "a", encoding="utf-8") as f:
#     f.write("起舞弄清影，\n")
#     f.write("何似在人间。\n")

# class ShoppingList:
#     def __init__(self, shopping_list):
#         self.shopping_list = shopping_list
#
#     def get_item_count(self):
#         return len(self.shopping_list)
#
#     def get_total_price(self):
#         toral_price = 0
#         for price in self.shopping_list.values():
#             toral_price += price
#         return toral_price

