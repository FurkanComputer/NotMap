import argparse
import socket
import sys
import time

# Bilinen portların zafiyet açıklamaları
PORT_EXPLANATIONS = {
    21: "FTP (Dosya Transferi): Şifrelenmemiş veri aktarımı. Brute-force riski taşır.",
    22: "SSH (Güvenli Kabuk): Uzaktan sunucu yönetimi. Güçsüz parola ve yetki yükseltme açıklarına dikkat edilmeli.",
    25: "SMTP (E-Posta): Mail sunucusu. Open-relay zafiyetleri kontrol edilmeli.",
    53: "DNS: IP-Domain dönüştürücü. DNS Amplification saldırılarında kullanılabilir.",
    80: "HTTP (Web Sunucusu): Şifrelenmemiş web trafiği. SQLi, XSS gibi web zafiyetleri taranmalı.",
    443: "HTTPS (Güvenli Web): SSL/TLS şifreli web trafiği.",
    3306: "MySQL Veritabanı: Dışarıya açıksa yetkisiz erişim ve veri sızıntısı riski oluşturur.",
    3389: "RDP (Uzak Masaüstü): Windows uzak bağlantısı. BlueKeep vb. kritik açıklara hedeftir."
}

# Taranacak en yaygın port listesi
COMMON_PORTS = {
    21: "ftp",
    22: "ssh",
    23: "telnet",
    25: "smtp",
    53: "dns",
    80: "http",
    110: "pop3",
    139: "netbios",
    443: "https",
    445: "microsoft-ds",
    1433: "mssql",
    3306: "mysql",
    3389: "ms-wbt-server",
    8080: "http-proxy"
}

def scan_port(ip, port, timeout=1.0):
    """Verilen IP ve Porta GERÇEK TCP bağlantısı dener."""
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        result = sock.connect_ex((ip, port))
        sock.close()
        return result == 0  # 0 dönerse port AÇIK demektir
    except Exception:
        return False

def main():
    parser = argparse.ArgumentParser(description="NotMap - Gerçek Port Tarayıcı")

    parser.add_argument("target", type=str, help="Hedef IP veya Alan Adı")
    parser.add_argument("--notmap-sv", action="store_true", help="Servis tespiti yap")
    parser.add_argument("--notmap-os", action="store_true", help="İşletim sistemi tahmini yap")
    parser.add_argument("--notmap-sn", action="store_true", help="Sadece ping/erişim kontrolü yap")
    parser.add_argument("--notmap-open", action="store_true", help="Sadece açık portları göster")
    parser.add_argument("--notmap-explain", action="store_true", help="Açık portların risklerini açıkla")
    parser.add_argument("--notmap-verbose", action="store_true", help="Detaylı çıktı ver")

    args = parser.parse_args()

    # 1. Alan Adını (Domain) IP Adresine Çevirme
    try:
        target_ip = socket.gethostbyname(args.target)
    except socket.gaierror:
        print(f"[!] Hata: '{args.target}' alan adı veya IP adresi çözümlenemedi!")
        sys.exit(1)

    print(f"[*] NotMap Taraması Başlatıldı -> Hedef: {args.target} ({target_ip})")
    print("-" * 60)

    # Ping / Erişim kontrolü
    if args.notmap_sn:
        print(f"[+] {target_ip} adresine bağlantı testi yapılıyor...")
        if scan_port(target_ip, 80) or scan_port(target_ip, 443) or scan_port(target_ip, 22):
            print(f"[+] Hedef Cihaz ({target_ip}): AKTİF (UP)")
        else:
            print(f"[!] Hedef Cihaz ({target_ip}): Yanıt vermiyor veya kapalı olabilir.")
        return

    open_ports = []

    print(f"{'PORT':<10} {'DURUM':<10} {'SERVİS':<15} {'VERSİYON'}")
    print("-" * 60)

    # 2. Gerçek Ağ Taraması DÖNGÜSÜ
    for port, service in COMMON_PORTS.items():
        if args.notmap_verbose:
            print(f"[...] Port {port} taranıyor...")

        is_open = scan_port(target_ip, port)

        if is_open:
            open_ports.append(port)
            version_str = "Tespit Edildi" if args.notmap_sv else "-"
            print(f"{str(port) + '/tcp':<10} {'open':<10} {service:<15} {version_str}")
        else:
            if not args.notmap_open:
                print(f"{str(port) + '/tcp':<10} {'closed':<10} {service:<15} -")

    # 3. Zafiyet / Amac Açıklamaları
    if args.notmap_explain and open_ports:
        print("\n" + "=" * 60)
        print(" AÇIK PORTLARIN AMACI VE GÜVENLİK RİSKLERİ (--notmap-explain)")
        print("=" * 60)
        for p in open_ports:
            info = PORT_EXPLANATIONS.get(p, "Bu port için özel risk tanımı yok.")
            print(f"[!] Port {p}: {info}")
    elif args.notmap_explain and not open_ports:
        print("\n[!] Hiç açık port bulunamadığı için açıklama oluşturulmadı.")

if __name__ == "__main__":
    main()