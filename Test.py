import openai

# 设置API密钥
openai.api_key = "your-api-key"


def generate_amazon_title(product_description):
    # 使用OpenAI的ChatCompletion API生成标题
    response = openai.ChatCompletion.create(
        model="gpt-4",  # 使用 GPT-4 模型（你可以根据需要选择 gpt-3.5-turbo）
        messages=[
            {"role": "system",
             "content": "You are a helpful assistant specialized in generating optimized Amazon product titles."},
            {"role": "user",
             "content": f"Create an Amazon product title based on this description: {product_description}"}
        ],
        max_tokens=60,  # 控制生成的标题长度
        temperature=0.7  # 控制生成结果的随机性
    )

    # 提取生成的标题
    title = response['choices'][0]['message']['content'].strip()
    return title


# 示例：生成亚马逊标题
product_description = "This is a lightweight, waterproof hiking backpack with a 30L capacity, suitable for both men and women, designed for outdoor activities such as camping and trekking."
title = generate_amazon_title(product_description)
print("Generated Amazon Title:", title)
