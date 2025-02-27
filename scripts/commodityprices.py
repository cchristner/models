import requests
from bs4 import BeautifulSoup

def scrape_agweb_futures():
    # URL for AgWeb futures page
    url = "https://www.agweb.com/markets/futures"
    
    # Set a user agent to mimic a browser
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Referer': 'https://www.agweb.com/',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1'
    }
    
    try:
        # Make the request
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        # Parse the HTML content
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Find all panels with class "panel panel-default"
        panels = soup.find_all('div', class_='panel panel-default')
        
        if not panels:
            print("No panels found. The website structure might have changed.")
            return
        
        # Process each panel
        for i, panel in enumerate(panels, 1):
            print(f"\n--- Panel {i} ---")
            
            # Get panel heading
            panel_heading = panel.find('div', class_='panel-heading')
            if panel_heading:
                print("Panel Heading:")
                print(panel_heading.get_text().strip())
            else:
                print("No panel heading found for this panel.")
            
            # Get panel body
            panel_body = panel.find('div', class_='panel-body')
            if panel_body:
                print("\nPanel Body:")
                print(panel_body.get_text().strip())
            else:
                print("\nNo panel body found for this panel.")
                
    except requests.exceptions.RequestException as e:
        print(f"Error making request: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    scrape_agweb_futures()