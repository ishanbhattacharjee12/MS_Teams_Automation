import requests

WEBHOOK_URL = "https://defaultff8e607edc6b4adf91383565a0f388.17.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/23/workflows/de8f7d403de64672abadceae1788d217/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=nIqbqt6ktEdyQwH8-D_BOaCFLuUenRyMK27xCR52bGs"

payload = {
    "text": " Test message from Python!"
}

try:
    response = requests.post(
        WEBHOOK_URL,
        json=payload,
        timeout=30
    )

    print("HTTP Status:", response.status_code)
    print("Response:", response.text)

    if response.ok:
        print(" Webhook request succeeded!")
    else:
        print(" Webhook request failed.")

except requests.RequestException as e:
    print(" Request error:", e)