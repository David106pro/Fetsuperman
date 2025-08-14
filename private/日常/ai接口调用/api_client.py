import requests
import json
import time
from datetime import datetime
import hmac
import hashlib
import base64

def get_auth_header(secret_id, secret_key, host, timestamp, payload):
    """生成腾讯云 API 认证头"""
    date = datetime.utcfromtimestamp(int(timestamp)).strftime('%Y-%m-%d')
    
    # 1. 拼接规范请求串
    http_request_method = "POST"
    canonical_uri = "/"
    canonical_querystring = ""
    
    # 规范化 Headers
    canonical_headers = (
        "content-type:application/json\n"
        f"host:{host}\n"
        "x-tc-action:chatcompletions\n"
    )
    signed_headers = "content-type;host;x-tc-action"
    
    # 请求体 Hash
    payload_hash = hashlib.sha256(json.dumps(payload).encode('utf-8')).hexdigest()
    
    # 2. 拼接待签名字符串
    canonical_request = (
        f"{http_request_method}\n"
        f"{canonical_uri}\n"
        f"{canonical_querystring}\n"
        f"{canonical_headers}\n"
        f"{signed_headers}\n"
        f"{payload_hash}"
    )
    
    # 3. 计算签名
    algorithm = "TC3-HMAC-SHA256"
    credential_scope = f"{date}/hunyuan/tc3_request"
    string_to_sign = (
        f"{algorithm}\n"
        f"{timestamp}\n"
        f"{credential_scope}\n"
        f"{hashlib.sha256(canonical_request.encode('utf-8')).hexdigest()}"
    )
    
    # 4. 计算签名
    def sign(key, msg):
        return hmac.new(key, msg.encode('utf-8'), hashlib.sha256).digest()
    
    secret_date = sign(("TC3" + secret_key).encode('utf-8'), date)
    secret_service = sign(secret_date, "hunyuan")
    secret_signing = sign(secret_service, "tc3_request")
    signature = hmac.new(secret_signing, string_to_sign.encode('utf-8'), hashlib.sha256).hexdigest()
    
    # 5. 组装 Authorization
    authorization = (
        f"{algorithm} "
        f"Credential={secret_id}/{credential_scope}, "
        f"SignedHeaders={signed_headers}, "
        f"Signature={signature}"
    )
    
    return authorization

def shorten_text(text, max_length=100, mode="normal"):
    """
    调用腾讯混元大模型API将文本缩短
    mode: 
        - normal: 普通简化模式
        - title: 生成标题模式，10-15字的吸引人标题
    """
    try:
        secret_id = "AKIDFEMTUM6odGiGi3HJSSJ8xhIgroW6RZT8"
        secret_key = "ayRxX5hGm8hDbmnfO1XryyoW3SineYUH"
        
        host = "hunyuan.tencentcloudapi.com"
        timestamp = str(int(time.time()))
        
        # 根据模式选择不同的提示语
        if mode == "title":
            prompt = f"请将以下文本提炼成一个10-15字的标题，要求：突出核心信息，吸引眼球，让人想进一步了解：\n\n{text}"
        else:
            prompt = f"请将以下文本缩短到{max_length}字以内，保留核心内容：\n\n{text}"
        
        payload = {
            "Model": "hunyuan-large",
            "Messages": [
                {
                    "Role": "user",
                    "Content": prompt
                }
            ],
            "Temperature": 0.7,
            "TopP": 0.7,
            "Stream": False
        }
        
        headers = {
            "Content-Type": "application/json",
            "Host": host,
            "X-TC-Action": "ChatCompletions",
            "X-TC-Version": "2023-09-01",
            "X-TC-Timestamp": timestamp,
            "X-TC-Region": "",
            "Authorization": get_auth_header(secret_id, secret_key, host, timestamp, payload)
        }

        response = requests.post(
            "https://hunyuan.tencentcloudapi.com/",
            headers=headers,
            json=payload,
            timeout=30
        )
        response.raise_for_status()
        result = response.json()
        
        if 'Response' in result and 'Choices' in result['Response']:
            return result['Response']['Choices'][0]['Message']['Content'].strip()
        else:
            raise Exception(f"API返回格式错误: {json.dumps(result)}")
            
    except requests.exceptions.RequestException as e:
        raise Exception(f"API请求失败: {str(e)}")
    except json.JSONDecodeError as e:
        raise Exception(f"API返回解析错误: {str(e)}")
    except Exception as e:
        raise Exception(f"处理失败: {str(e)}")
    finally:
        time.sleep(1)  # 移到 finally 块中 