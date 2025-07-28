from msal import PublicClientApplication
import requests, pandas as pd
import os, time
from tqdm import tqdm  # For progress bar
from dotenv import load_dotenv

load_dotenv() 

print("🚀 Script started")

# Auth setup (unchanged)
client_id = os.getenv('CLIENT_ID')
tenant_id = os.getenv('TENANT_ID')
authority = f"{os.getenv('AUTHORITY')}/{tenant_id}"
scopes = os.getenv('SCOPES').split()  

app = PublicClientApplication(client_id, authority=authority)
flow = app.initiate_device_flow(scopes=scopes)

if 'user_code' not in flow:
    raise ValueError("Device code flow failed")

print(f"\n👉 Visit {flow['verification_uri']} and enter code: {flow['user_code']}")
result = app.acquire_token_by_device_flow(flow)

if 'access_token' in result:
    print("✅ Authenticated successfully")
    headers = {'Authorization': f"Bearer {result['access_token']}"}
    
    # Get all users (with pagination)
    print("⏳ Fetching users...")
    users = []
    url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,assignedLicenses,department"
    while url:
        response = requests.get(url, headers=headers).json()
        users.extend(response.get('value', []))
        url = response.get('@odata.nextLink')
    print(f"👥 Found {len(users)} users")

    # Get license SKU names
    sku_map = {}
    sku_resp = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus", headers=headers).json()
    for sku in sku_resp.get('value', []):
        sku_map[sku['skuId']] = sku['skuPartNumber']

    # Get mailbox info for each user
    print("⏳ Fetching mailbox sizes (this may take a few minutes)...")
    report = []
    
    for user in tqdm(users):  # Progress bar
        upn = user.get('userPrincipalName')
        display = user.get('displayName')
        user_id = user.get('id')
        department = user.get('department')
        
        # Get licenses
        sku_ids = [l['skuId'] for l in user.get('assignedLicenses', [])]
        license_names = [sku_map.get(s, s) for s in sku_ids]
        
        # Get mailbox settings
        mailbox_size = None
        try:
            settings = requests.get(
                f"https://graph.microsoft.com/v1.0/users/{user_id}/mailboxSettings",
                headers=headers,
                timeout=10
            ).json()
            
            if 'storageQuota' in settings:
                used_bytes = settings['storageQuota'].get('used')
                if used_bytes:
                    mailbox_size = round(int(used_bytes) / (1024 * 1024), 2)  # Convert to MB
        except Exception as e:
            pass  # Silently skip errors

        report.append({
            "DisplayName": display,
            "UserPrincipal": upn,
            "Licenses": ", ".join(license_names),
            "MailboxSizeMB": mailbox_size,
            "Department": department
        })
        time.sleep(0.1)  # Rate limiting

    # Save report
    random_suffix = os.urandom(4).hex()
    df = pd.DataFrame(report)
    output_path = os.path.join(os.getcwd(), f"m365_users_mailbox_{random_suffix}.csv")
    df.to_csv(output_path, index=False)
    print(f"📄 Report saved to: {output_path}")

else:
    print("❌ Authentication failed")
