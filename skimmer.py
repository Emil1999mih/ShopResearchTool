import requests
import json
#alte tipuri de site-uri(wix, wordpress. etc.)
def extract_data(url):   #functia care incarca si extrage datele din products.json

    urlfinalizat =url+"/products.json"
    r = requests.get(urlfinalizat)
    data = r.json()
    product_list = []

    for item in data['products']:
        title = item['title']
        handle = item['handle']

        for image in item['images']:
            try:
                imagesrc = image['src']
            except:
                imagesrc = 'None'

        for variant in item['variants']:
            price = variant['price']
            sku = variant['sku']
            grams = variant['grams']
            available = variant['available']

        product = {
            'title': title,
            'handle': handle,
            'price': price,
            'sku': sku,
            'image': imagesrc,
            'available': available
        }
        product_list.append(product)

    return product_list


