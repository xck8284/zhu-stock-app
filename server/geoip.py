import requests


def lookup_ip_info(ip: str) -> dict:
    if not ip or ip in ["127.0.0.1", "localhost", "::1"]:
        return {
            "ip": ip,
            "country": "Local",
            "region": "Local",
            "city": "Local",
            "loc": "",
            "org": "",
            "timezone": "",
        }

    try:
        r = requests.get(f"https://ipinfo.io/{ip}/json", timeout=5)
        data = r.json()
        return {
            "ip": data.get("ip", ip),
            "country": data.get("country", ""),
            "region": data.get("region", ""),
            "city": data.get("city", ""),
            "loc": data.get("loc", ""),
            "org": data.get("org", ""),
            "timezone": data.get("timezone", ""),
        }
    except Exception:
        return {
            "ip": ip,
            "country": "",
            "region": "",
            "city": "",
            "loc": "",
            "org": "",
            "timezone": "",
        }