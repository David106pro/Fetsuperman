import re
text = "野孩子,王俊凯,邓家佳,陈永胜,殷若昕,选座购票,剧情介绍"
result = re.sub(r'^.*?,|,选座购票.*$', '', text)
print(result)