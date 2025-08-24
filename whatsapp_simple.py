import subprocess
import psutil
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

class WhatsAppAutomation:
    def __init__(self):
        """WhatsApp Blasting Automation"""
        self.driver = None
        self.wait = None
        self.debug_port = 9222
        self.chrome_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        self.temp_profile = r"C:\temp\chrome_debug"
        self.main_page_url = "http://192.168.18.32:3000/blasting_monitor.html"
        
        self.main_tab = None
        self.whatsapp_tab = None
        
        print("🚀 WhatsApp Blasting Automation")
        print("=" * 55)
    
    def setup_chrome(self):
        """Setup Chrome dengan debugging"""
        print("🔄 Setup Chrome...")
        
        # Force close Chrome
        try:
            subprocess.run(["taskkill", "/f", "/im", "chrome.exe"], 
                          capture_output=True, check=False)
            time.sleep(3)
        except:
            pass
        
        # Create temp directory
        if not os.path.exists(self.temp_profile):
            os.makedirs(self.temp_profile)
        
        # Start Chrome dengan debugging
        chrome_cmd = [
            self.chrome_path,
            f"--remote-debugging-port={self.debug_port}",
            f"--user-data-dir={self.temp_profile}",
            "--no-first-run",
            "--start-maximized"
        ]
        
        subprocess.Popen(chrome_cmd)
        time.sleep(5)
        
        # Connect ke Chrome
        options = Options()
        options.add_experimental_option("debuggerAddress", f"localhost:{self.debug_port}")
        
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 15)
        
        print("   ✅ Chrome ready")
        return True
    
    def open_pages(self):
        """Buka halaman utama dan WhatsApp"""
        print("📄 Membuka halaman...")
        
        # Buka halaman utama
        self.driver.get(self.main_page_url)
        time.sleep(3)
        self.main_tab = self.driver.current_window_handle
        
        # Buka WhatsApp Web tab
        self.driver.execute_script("window.open('https://web.whatsapp.com/', '_blank');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.whatsapp_tab = self.driver.current_window_handle
        time.sleep(5)
        
        # Kembali ke main tab
        self.driver.switch_to.window(self.main_tab)
        
        print("   ✅ Pages ready")
        return True
    
    def parse_input(self, user_input):
        """Parse input user untuk nomor baris"""
        print(f"🔢 Parsing: '{user_input}'")
        
        try:
            if user_input.lower() == 'all':
                # Generate semua baris (1-565)
                return list(range(1, 566))
            
            row_numbers = set()
            parts = user_input.replace(" ", "").split(",")
            
            for part in parts:
                if "-" in part:
                    # Range format
                    start, end = part.split("-")
                    start_num = int(start)
                    end_num = int(end)
                    
                    for i in range(start_num, end_num + 1):
                        row_numbers.add(i)
                else:
                    # Single number
                    row_numbers.add(int(part))
            
            result = sorted(list(row_numbers))
            print(f"   ✅ Baris: {result}")
            return result
            
        except Exception as e:
            print(f"   ❌ Error: {e}")
            return []
    
    def get_row_input(self):
        """Input nomor baris dari user"""
        print("\n" + "="*60)
        print("📋 PILIH BARIS YANG INGIN DIPROSES")
        print("="*60)
        print("Format:")
        print("  • 1-10     → Baris 1 sampai 10")
        print("  • 1,5,10   → Baris 1, 5, dan 10")
        print("  • 1-3,10   → Baris 1-3 dan 10")
        print("  • all      → Semua baris (1-565)")
        print()
        
        while True:
            user_input = input("Input nomor baris: ").strip()
            
            if not user_input:
                print("❌ Input kosong!")
                continue
            
            if user_input.lower() == 'all':
                print("⚠️  Mode ALL - 565 baris (~7+ jam)")
                confirm = input("Lanjutkan? (yes/no): ")
                if confirm.lower() != 'yes':
                    continue
                return list(range(1, 566))
            
            rows = self.parse_input(user_input)
            if not rows:
                print("❌ Input tidak valid!")
                continue
            
            print(f"✅ {len(rows)} baris akan diproses")
            confirm = input("Lanjutkan? (y/n): ")
            if confirm.lower() == 'y':
                return rows
    
    def click_copy(self, row_num):
        """Klik Copy button"""
        try:
            selectors = [
                f"//table//tr[{row_num}]//button[contains(text(), 'Copy')]",
                f"(//button[contains(text(), 'Copy')])[{row_num}]"
            ]
            
            for selector in selectors:
                try:
                    button = self.driver.find_element(By.XPATH, selector)
                    if button.is_displayed():
                        button.click()
                        time.sleep(2)
                        return True
                except:
                    continue
            return False
        except:
            return False
    
    def click_wa(self, row_num):
        """Klik WA button"""
        try:
            selectors = [
                f"//table//tr[{row_num}]//a[contains(@class, 'btn-wa')]",
                f"(//a[contains(@class, 'btn-wa')])[{row_num}]"
            ]
            
            for selector in selectors:
                try:
                    link = self.driver.find_element(By.XPATH, selector)
                    if link.is_displayed():
                        link.click()
                        time.sleep(3)
                        return True
                except:
                    continue
            return False
        except:
            return False
    
    def handle_whatsapp(self):
        """Handle WhatsApp flow"""
        try:
            # Switch ke tab terbaru
            all_tabs = self.driver.window_handles
            if len(all_tabs) > 2:
                self.driver.switch_to.window(all_tabs[-1])
            
            time.sleep(3)
            
            # Klik "Lanjut ke Chat"
            print("      🔍 Mencari 'Lanjut ke Chat'...")
            lanjut_selectors = [
                "//span[contains(text(), 'Lanjut ke Chat')]",
                "//*[contains(text(), 'Lanjut ke Chat')]"
            ]
            
            lanjut_clicked = False
            for selector in lanjut_selectors:
                try:
                    button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    button.click()
                    print("      ✅ 'Lanjut ke Chat' diklik")
                    time.sleep(3)
                    lanjut_clicked = True
                    break
                except:
                    continue
            
            if not lanjut_clicked:
                print("      ⚠️ 'Lanjut ke Chat' tidak ditemukan, lanjut...")
            
            # Klik "gunakan WhatsApp Web"
            print("      🔍 Mencari 'gunakan WhatsApp Web'...")
            web_selectors = [
                "//span[contains(text(), 'gunakan WhatsApp Web')]",
                "//a[contains(text(), 'gunakan WhatsApp Web')]",
                "//a[contains(@href, 'web.whatsapp.com')]"
            ]
            
            web_clicked = False
            for selector in web_selectors:
                try:
                    link = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    link.click()
                    print("      ✅ 'gunakan WhatsApp Web' diklik")
                    time.sleep(5)  # Tunggu redirect
                    web_clicked = True
                    break
                except:
                    continue
            
            if not web_clicked:
                print("      ⚠️ 'gunakan WhatsApp Web' tidak ditemukan, lanjut...")
            
            # Tunggu WhatsApp Web benar-benar loaded
            print("      ⏳ Menunggu WhatsApp Web loaded...")
            time.sleep(8)  # Extra wait untuk ensure loaded
            
            # Cari message chat box yang BENAR berdasarkan HTML yang Anda tunjukkan
            print("      🔍 Mencari message chat box...")
            
            # Selector yang SANGAT SPESIFIK berdasarkan HTML Anda
            message_selectors = [
                # Berdasarkan HTML yang Anda tunjukkan - PALING AKURAT
                "//div[@aria-label='Ketik pesan' and @role='textbox' and @contenteditable='true' and @data-tab='10']",
                
                # Alternatif dengan aria-placeholder
                "//div[@aria-placeholder='Ketik pesan' and @contenteditable='true']",
                
                # Kombinasi attributes
                "//div[@role='textbox' and @data-tab='10' and contains(@aria-label, 'Ketik')]",
                
                # Berdasarkan class dan attributes
                "//div[@contenteditable='true' and @data-tab='10' and contains(@class, 'selectable-text')]",
                
                # Footer area dengan attributes spesifik  
                "//footer//div[@role='textbox' and @contenteditable='true']",
                
                # Fallback yang lebih umum tapi dengan filter
                "//div[@contenteditable='true' and @role='textbox']"
            ]
            
            message_box = None
            for i, selector in enumerate(message_selectors, 1):
                try:
                    print(f"         Selector {i}: {selector[:60]}...")
                    
                    # Tunggu element muncul dengan timeout lebih lama
                    elements = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_all_elements_located((By.XPATH, selector))
                    )
                    
                    for element in elements:
                        if element.is_displayed():
                            # Validasi extra: pastikan ini message box bukan search box
                            aria_label = element.get_attribute('aria-label') or ''
                            aria_placeholder = element.get_attribute('aria-placeholder') or ''
                            data_tab = element.get_attribute('data-tab') or ''
                            
                            print(f"            Aria-label: {aria_label}")
                            print(f"            Aria-placeholder: {aria_placeholder}")
                            print(f"            Data-tab: {data_tab}")
                            
                            # Pastikan ini message box (bukan search)
                            if ('ketik pesan' in aria_label.lower() or 
                                'ketik pesan' in aria_placeholder.lower() or
                                'type a message' in aria_label.lower() or
                                data_tab == '10'):
                                
                                print(f"         ✅ Message box BENAR ditemukan!")
                                message_box = element
                                break
                            else:
                                print(f"         ❌ Bukan message box: {aria_label}")
                    
                    if message_box:
                        break
                        
                except Exception as e:
                    print(f"         ❌ Selector {i} error: {str(e)[:50]}")
                    continue
            
            if message_box:
                print("      ✅ Message chat box siap!")
                return message_box
            else:
                print("      ❌ Message chat box tidak ditemukan")
                
                # Debug lebih detail
                try:
                    print("      🔍 Debug semua textbox:")
                    all_textboxes = self.driver.find_elements(By.XPATH, "//div[@contenteditable='true']")
                    
                    for i, box in enumerate(all_textboxes):
                        if box.is_displayed():
                            aria_label = box.get_attribute('aria-label') or 'No label'
                            aria_placeholder = box.get_attribute('aria-placeholder') or 'No placeholder'
                            data_tab = box.get_attribute('data-tab') or 'No data-tab'
                            
                            print(f"         {i+1}. Label: {aria_label}")
                            print(f"            Placeholder: {aria_placeholder}")
                            print(f"            Data-tab: {data_tab}")
                            print(f"            ---")
                except:
                    pass
                
                return None
            
        except Exception as e:
            print(f"      ❌ Error handle WhatsApp: {e}")
            return None
    
    def send_message(self, chat_box):
        """Paste dan kirim pesan"""
        try:
            chat_box.click()
            time.sleep(1)
            chat_box.send_keys(Keys.CONTROL + 'v')
            time.sleep(2)
            chat_box.send_keys(Keys.ENTER)
            time.sleep(3)
            return True
        except:
            return False
    
    def update_status(self, row_num, status):
        """Update status button"""
        try:
            # Kembali ke main tab
            self.driver.switch_to.window(self.main_tab)
            time.sleep(1)
            
            selector = f"//table//tr[{row_num}]//button[contains(text(), '{status}')]"
            button = self.driver.find_element(By.XPATH, selector)
            button.click()
            time.sleep(1)
            return True
        except:
            return False
    
    def close_wa_tab(self):
        """Tutup tab WhatsApp dan kembali ke main"""
        try:
            self.driver.close()
            remaining_tabs = self.driver.window_handles
            if remaining_tabs:
                self.driver.switch_to.window(remaining_tabs[0])
                self.main_tab = remaining_tabs[0]
            return True
        except:
            return False
    
    def process_row(self, row_num):
        """Proses satu baris"""
        print(f"   🔄 Proses baris {row_num}")
        
        # Pastikan di main tab
        self.driver.switch_to.window(self.main_tab)
        
        # 1. Copy
        if not self.click_copy(row_num):
            print(f"      ❌ Copy gagal")
            self.update_status(row_num, "Gagal")
            return False
        print(f"      ✅ Copy berhasil")
        
        # 2. WA
        if not self.click_wa(row_num):
            print(f"      ❌ WA gagal")
            self.update_status(row_num, "Gagal")
            return False
        print(f"      ✅ WA berhasil")
        
        # 3. Handle WhatsApp dengan retry dan timeout yang lebih baik
        print(f"      📱 Handling WhatsApp...")
        
        max_retries = 3
        chat_box = None
        
        for attempt in range(max_retries):
            print(f"         Attempt {attempt + 1}/{max_retries}")
            
            chat_box = self.handle_whatsapp()
            
            if chat_box:
                print(f"      ✅ WhatsApp berhasil (attempt {attempt + 1})")
                break
            else:
                print(f"      ⚠️ WhatsApp attempt {attempt + 1} gagal")
                if attempt < max_retries - 1:
                    print(f"         Retry dalam 3 detik...")
                    time.sleep(3)
        
        if not chat_box:
            print(f"      ❌ WhatsApp gagal setelah {max_retries} attempts")
            try:
                self.close_wa_tab()
            except:
                pass
            self.update_status(row_num, "Gagal")
            return False
        
        # 4. Send message
        print(f"      📤 Mengirim pesan...")
        if not self.send_message(chat_box):
            print(f"      ❌ Send gagal")
            try:
                self.close_wa_tab()
            except:
                pass
            self.update_status(row_num, "Gagal")
            return False
        print(f"      ✅ Send berhasil")
        
        # 5. Tunggu sebentar untuk memastikan pesan terkirim
        print(f"      ⏳ Tunggu konfirmasi kirim...")
        time.sleep(3)
        
        # 6. Close dan update
        print(f"      🔚 Tutup tab WhatsApp...")
        if not self.close_wa_tab():
            print(f"      ⚠️ Gagal tutup tab, tapi lanjut...")
        
        # 7. Update status di main tab
        print(f"      ✅ Update status...")
        time.sleep(1)  # Pastikan sudah di main tab
        self.update_status(row_num, "Terkirim")
        
        print(f"      ✅ Baris {row_num} selesai")
        return True
    
    def run_automation(self, row_numbers):
        """Main automation loop"""
        total = len(row_numbers)
        success = 0
        failed = 0
        
        print(f"🚀 Mulai automasi {total} baris")
        
        for i, row in enumerate(row_numbers, 1):
            print(f"\n📍 {i}/{total} - Baris {row}")
            
            try:
                if self.process_row(row):
                    success += 1
                else:
                    failed += 1
            except Exception as e:
                print(f"   ❌ Error: {e}")
                failed += 1
            
            # Delay
            if i < total:
                print(f"   ⏳ Delay 3 detik...")
                time.sleep(3)
        
        print(f"\n📊 HASIL:")
        print(f"✅ Berhasil: {success}")
        print(f"❌ Gagal: {failed}")
        print(f"📈 Rate: {success/total*100:.1f}%")
        
        return success, failed
    
    def run(self):
        """Main function"""
        try:
            # Setup
            if not self.setup_chrome():
                return False
            
            if not self.open_pages():
                return False
            
            # Konfirmasi WhatsApp
            print("📱 Pastikan WhatsApp Web sudah login...")
            input("Tekan Enter jika siap: ")
            
            # Input baris
            row_numbers = self.get_row_input()
            if not row_numbers:
                return False
            
            # Run automation
            self.run_automation(row_numbers)
            
            print("✅ Selesai!")
            return True
            
        except Exception as e:
            print(f"❌ Error: {e}")
            return False
    
    def close(self):
        """Tutup browser"""
        if self.driver:
            self.driver.quit()

def main():
    automation = WhatsAppAutomation()
    
    try:
        automation.run()
        
        choice = input("\nTutup browser? (y/n): ")
        if choice.lower() == 'y':
            automation.close()
        
    except KeyboardInterrupt:
        print("\n🛑 Dihentikan")
        automation.close()
    except Exception as e:
        print(f"\n❌ Error: {e}")
        automation.close()

if __name__ == "__main__":
    main()
    
# sdhjfsdasfe