import base64


def image_to_py(image_path):
    with open(image_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode('utf-8')

    with open('ico.txt', "w") as f:
        f.write(encoded)
    return encoded

# Использование:
print(image_to_py("MTC_Logo_CMYK.png"))