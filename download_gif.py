import urllib.request
import re
import sys
import os

url = "https://tenor.com/hcTSI6uzhjH.gif"
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

print(f"Fetching page: {url}")
req = urllib.request.Request(url, headers=headers)
try:
    with urllib.request.urlopen(req) as response:
        html = response.read().decode('utf-8', errors='ignore')
except Exception as e:
    print(f"Error fetching URL: {e}")
    sys.exit(1)

# Find the main GIF URL from meta tags
# Format could be: <meta class="dynamic" property="og:image" content="...">
og_image_match = re.search(r'<meta[^>]*property=["\']og:image["\'][^>]*content=["\']([^"\']+)["\']', html)
if not og_image_match:
    # Alternative: content comes before property
    og_image_match = re.search(r'<meta[^>]*content=["\']([^"\']+)["\'][^>]*property=["\']og:image["\']', html)

if og_image_match:
    target_url = og_image_match.group(1).replace('&amp;', '&').replace('\\u002F', '/')
    print(f"Found og:image: {target_url}")
else:
    # Try twitter:image
    twitter_image_match = re.search(r'<meta[^>]*name=["\']twitter:image["\'][^>]*content=["\']([^"\']+)["\']', html)
    if not twitter_image_match:
        twitter_image_match = re.search(r'<meta[^>]*content=["\']([^"\']+)["\'][^>]*name=["\']twitter:image["\']', html)
        
    if twitter_image_match:
        target_url = twitter_image_match.group(1).replace('&amp;', '&').replace('\\u002F', '/')
        print(f"Found twitter:image: {target_url}")
    else:
        # Fallback to the first media URL ending in .gif
        matches = re.findall(r'https://media1?\.tenor\.com/[^"\'>]+\.gif', html)
        if matches:
            target_url = matches[0].replace('&amp;', '&').replace('\\u002F', '/')
            print(f"Fallback to first GIF: {target_url}")
        else:
            target_url = None

if target_url:
    # Tenor sometimes serves a low resolution one, but media1.tenor.com/m/... is usually the original or optimized one
    print(f"Downloading main GIF from: {target_url}")
    output_filename = "tidal_wave.gif"
    
    req_gif = urllib.request.Request(target_url, headers=headers)
    try:
        with urllib.request.urlopen(req_gif) as response_gif, open(output_filename, 'wb') as out_file:
            out_file.write(response_gif.read())
        print(f"Successfully downloaded to {output_filename}!")
    except Exception as e:
        print(f"Error downloading GIF: {e}")
else:
    print("Could not find the main GIF on the page.")
