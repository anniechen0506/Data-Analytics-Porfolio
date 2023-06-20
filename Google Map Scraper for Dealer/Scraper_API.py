import requests
import json

url = 'https://maps.googleapis.com/maps/api/place/details/json'
params = {
    'place_id': 'ChIJ7yui6X-sQjQRkmX-NEcZYbM', # Replace with the unique ID of the Audi dealer
    'fields': 'reviews',
    'key': '# Replace with your own API key', 
    'max_results' : 100 
}

# Send a GET request to the API endpoint
response = requests.get(url,params=params)
data = response.json()

if 'result' in data:
    reviews = data['result'].get('reviews', [])
    next_page_token = data.get('next_page_token')
    
    while next_page_token:
        params['pagetoken'] = next_page_token
        response = requests.get(url, params=params)
        data = response.json()
        reviews.extend(data['result'].get('reviews', []))
        next_page_token = data.get('next_page_token')
        
for review in reviews:
    print(f"Rating: {review['rating']}")
    print(f"Text: {review['text']}\n")
    
# # Extract the reviews from the response
# if 'result' in response.json():
#     reviews = response.json()['result'].get('reviews', [])
# else:
#     reviews = []

# # Print the text and rating of each review
# for review in reviews:
#     print(f"Rating: {review['rating']}")
#     print(f"Text: {review['text']}\n")
