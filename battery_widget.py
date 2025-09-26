import tkinter as tk
from tkinter import messagebox
import subprocess, threading, json, os, re
from datetime import datetime

class BatteryHealthGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ASUS ZenBook 14 - Battery Health Monitor")
        self.root.geometry("550x700")
        self.root.configure(bg='#1a1a1a')
        self.root.resizable(True, True)

        self.saved_data = self.load_saved_data()
        self.current_data = None

        self.setup_ui()
        self.attempt_auto_detection()

    # -------------------- UI SETUP --------------------
    def setup_ui(self):
        header = tk.Label(self.root,text="🔋 ASUS ZenBook 14 Battery Health",
                          bg='#1a1a1a',fg='white',font=('Arial',16,'bold'))
        header.pack(pady=(10,0))
        tk.Label(self.root,text="Monitor your battery health and performance",
                 bg='#1a1a1a',fg='#aaaaaa',font=('Arial',10)).pack(pady=(0,10))

        self.status_label = tk.Label(self.root,text="Ready to scan battery...",
                                     bg='#333333',fg='#ffff00',font=('Arial',10,'bold'))
        self.status_label.pack(fill='x',pady=(0,10))

        # Battery Info
        self.info_frame = tk.Frame(self.root,bg='#2d2d2d',bd=1,relief='ridge')
        self.info_frame.pack(fill='x',padx=20,pady=10)
        self.charge_label  = self.row(self.info_frame,"⚡ Current Charge","--%")
        self.state_label   = self.row(self.info_frame,"📊 Status","--")
        self.power_label   = self.row(self.info_frame,"🔌 Power Source","--")
        self.design_label  = self.row(self.info_frame,"🏭 Design Capacity","-- mWh")
        self.full_label    = self.row(self.info_frame,"🔋 Full Charge Capacity","-- mWh")
        self.cycle_label   = self.row(self.info_frame,"🔄 Cycle Count","--")
        self.health_label  = self.row(self.info_frame,"❤️ Battery Health","--%")

        self.health_canvas = tk.Canvas(self.info_frame,width=350,height=30,
                                       bg='#3a3a3a',highlightthickness=1,
                                       highlightbackground='#555555')
        self.health_canvas.pack(pady=10)
        self.assess_label = tk.Label(self.info_frame,text="Assessment: --",
                                     bg='#2d2d2d',fg='white',font=('Arial',10,'bold'))
        self.assess_label.pack()
        self.update_label = tk.Label(self.info_frame,text="Last updated: --",
                                     bg='#2d2d2d',fg='#888888',font=('Arial',8))
        self.update_label.pack(pady=5)

        # Buttons
        btn_frame = tk.Frame(self.root,bg='#1a1a1a')
        btn_frame.pack(pady=10)
        tk.Button(btn_frame,text="🔄 Auto Detect",command=self.attempt_auto_detection,
                  bg='#4CAF50',fg='white',font=('Arial',10),padx=12,pady=5).pack(side='left',padx=5)
        tk.Button(btn_frame,text="📊 Generate Report",command=self.generate_report,
                  bg='#FF9800',fg='white',font=('Arial',10),padx=12,pady=5).pack(side='left',padx=5)
        tk.Button(btn_frame,text="💾 Save",command=self.save_current_data,
                  bg='#9C27B0',fg='white',font=('Arial',10),padx=12,pady=5).pack(side='left',padx=5)
        tk.Button(btn_frame,text="🗑️ Clear",command=self.clear_data,
                  bg='#f44336',fg='white',font=('Arial',10),padx=12,pady=5).pack(side='left',padx=5)

        # Manual Entry (embedded)
        self.manual_frame = tk.LabelFrame(self.root,text="Manual Entry",
                                          bg='#2d2d2d',fg='white',font=('Arial',11,'bold'))
        self.manual_frame.pack(fill='x',padx=20,pady=10)
        self.entries = {}
        for lbl,key in [
            ("🔋 Design Capacity (mWh):","design"),
            ("⚡ Full Charge Capacity (mWh):","full"),
            ("🔄 Cycle Count:","cycle"),
            ("📊 Current Charge (%):","percent")
        ]:
            row = tk.Frame(self.manual_frame,bg='#2d2d2d'); row.pack(fill='x',padx=10,pady=4)
            tk.Label(row,text=lbl,bg='#2d2d2d',fg='#cccccc',width=22,anchor='w').pack(side='left')
            e = tk.Entry(row,width=15); e.pack(side='left'); self.entries[key]=e
        tk.Button(self.manual_frame,text="✅ Submit Manual Data",
                  command=self.submit_manual,bg='#2196F3',fg='white',
                  font=('Arial',10),padx=10,pady=5).pack(pady=10)

        if self.saved_data:
            self.update_display(self.saved_data)

    def row(self,parent,label,value):
        f = tk.Frame(parent,bg='#2d2d2d'); f.pack(fill='x',padx=10,pady=2)
        tk.Label(f,text=label,bg='#2d2d2d',fg='#cccccc',width=25,anchor='w').pack(side='left')
        val = tk.Label(f,text=value,bg='#2d2d2d',fg='white',anchor='w')
        val.pack(side='left')
        return val

    # -------------------- Data Logic --------------------
    def load_saved_data(self):
        if os.path.exists('battery_data.json'):
            try:
                with open('battery_data.json','r') as f: return json.load(f)
            except: return {}
        return {}

    def save_data(self,data):
        try:
            with open('battery_data.json','w') as f: json.dump(data,f,indent=2)
        except Exception:
            pass

    def attempt_auto_detection(self):
        self.status_label.config(text="🔄 Detecting battery info...",fg="#ffff00")
        threading.Thread(target=self.detect_thread,daemon=True).start()

    def detect_thread(self):
        # 1) Try WMI (PowerShell)
        info = self.try_wmi()
        # 2) If WMI returned but capacity is missing or zero, try powercfg report parsing
        if info:
            design = info.get('design_capacity')
            full = info.get('full_capacity')
            # consider missing or zero or None as invalid
            if not self._is_positive_number(design) or not self._is_positive_number(full):
                # try powercfg parsing fallback
                fallback = self.try_powercfg_report_parse()
                if fallback:
                    # update missing fields
                    if self._is_positive_number(fallback.get('design_capacity')):
                        info['design_capacity'] = fallback['design_capacity']
                    if self._is_positive_number(fallback.get('full_capacity')):
                        info['full_capacity'] = fallback['full_capacity']
                    # recalc health if possible
                    if self._is_positive_number(info.get('design_capacity')) and self._is_positive_number(info.get('full_capacity')):
                        info['health_percent'] = (info['full_capacity'] / info['design_capacity']) * 100
                        # cap to reasonable 0-100
                        info['health_percent'] = max(0.0, min(100.0, float(info['health_percent'])))
                    self.root.after(0, lambda: self.status_label.config(text="🔄 Used powercfg fallback to refine capacities", fg="#ffff00"))
            # ensure health exists
            if 'health_percent' not in info or info.get('health_percent') is None:
                d = info.get('design_capacity') or 0
                f = info.get('full_capacity') or 0
                info['health_percent'] = (f/d*100) if (d and f) else None
        else:
            # try simple percentage-only method
            info = self.try_simple()

        self.root.after(0, lambda: self.handle_result(info))

    def try_wmi(self):
        """Use WMI to get detailed info (Windows only)."""
        ps = (
            "$b=Get-WmiObject -Class Win32_Battery;"
            "if($b){@{Percent=$b.EstimatedChargeRemaining;"
            "DesignCapacity=$b.DesignCapacity;FullCapacity=$b.FullChargeCapacity;"
            "CycleCount=$b.CycleCount;Status=$b.BatteryStatus;Plugged=$b.PowerOnline}|ConvertTo-Json}"
        )
        try:
            r = subprocess.run(["powershell","-Command",ps],
                               capture_output=True,text=True,timeout=15)
            out = r.stdout.strip()
            if out:
                # Parse JSON safely and convert numeric strings to numbers when possible
                try:
                    d = json.loads(out)
                except Exception:
                    # Sometimes ConvertTo-Json returns arrays or single-line dictionaries; try to fix minor issues
                    try:
                        out_fixed = out.replace("\n"," ").strip()
                        d = json.loads(out_fixed)
                    except Exception:
                        return None
                return self.process_data(d)
        except Exception:
            return None

    def try_simple(self):
        try:
            r = subprocess.run(["powershell","-Command",
                                "(Get-WmiObject -Class Win32_Battery).EstimatedChargeRemaining"],
                               capture_output=True,text=True,timeout=10)
            if r.stdout and r.stdout.strip().isdigit():
                p = int(r.stdout.strip())
                return {'percent': p, 'design_capacity': None, 'full_capacity': None,
                        'cycle_count': None, 'status': 1, 'plugged': False,
                        'health_percent': None}
        except Exception:
            return None

    def process_data(self,d):
        def safe_int_from(obj, key):
            v = obj.get(key)
            if isinstance(v, (int, float)):
                return int(v)
            if isinstance(v, str):
                s = v.strip().replace(",", "")
                # sometimes JSON returns boolean/None, guard that:
                if s.isdigit():
                    return int(s)
            return None

        percent = safe_int_from(d, "Percent")
        design = safe_int_from(d, "DesignCapacity")
        full = safe_int_from(d, "FullCapacity") or safe_int_from(d, "FullChargeCapacity")
        cycle = safe_int_from(d, "CycleCount")
        status = d.get("Status") if isinstance(d.get("Status"), (int, float)) else None
        plugged = d.get("Plugged") if isinstance(d.get("Plugged"), bool) else None

        # health calculation if both capacities present and >0
        health = None
        if self._is_positive_number(design) and self._is_positive_number(full):
            health = (full / design) * 100
            health = float(max(0.0, min(100.0, health)))

        return {
            'percent': percent,
            'design_capacity': design,
            'full_capacity': full,
            'cycle_count': cycle,
            'status': status,
            'plugged': plugged,
            'health_percent': health
        }

    def try_powercfg_report_parse(self):
        """Generate battery-report.html and extract design/full capacity as fallback"""
        try:
            out_path = os.path.join(os.environ.get("USERPROFILE", "."), "battery-report.html")
            # Try to generate the battery report
            subprocess.run(["powercfg", "/batteryreport", "/output", out_path],
                           capture_output=True, text=True, timeout=30)

            if not os.path.exists(out_path):
                return None

            with open(out_path, "r", encoding="utf-8", errors="ignore") as f:
                html = f.read()

            # Try to find the 'Design Capacity' and 'Full Charge Capacity' values in mWh
            # The report may contain multiple occurrences; pick the first reasonable match.
            design_match = re.search(r"Design (?:Capacity|capacity).*?([0-9\.,]+)\s*mWh", html, re.IGNORECASE | re.DOTALL)
            full_match = re.search(r"Full (?:Charge )?Capacity.*?([0-9\.,]+)\s*mWh", html, re.IGNORECASE | re.DOTALL)

            # If direct matches failed, try slightly different patterns
            if not design_match:
                design_match = re.search(r"Installed batteries[\s\S]{0,200}Design.*?([0-9\.,]+)\s*mWh", html, re.IGNORECASE)
            if not full_match:
                full_match = re.search(r"Installed batteries[\s\S]{0,200}Full.*?([0-9\.,]+)\s*mWh", html, re.IGNORECASE)

            if design_match and full_match:
                d_str = design_match.group(1).replace(",", "").replace(".", "")
                f_str = full_match.group(1).replace(",", "").replace(".", "")
                # Some HTML may show values like "50,000" or "50000" or "50.000" — we removed punctuation above.
                try:
                    design_val = int(d_str)
                    full_val = int(f_str)
                except Exception:
                    return None

                # calculate health and clamp
                health_val = float(full_val) / float(design_val) * 100.0 if design_val and full_val else None
                if health_val is not None:
                    health_val = max(0.0, min(100.0, health_val))

                return {
                    'design_capacity': design_val,
                    'full_capacity': full_val,
                    'health_percent': health_val
                }

        except Exception:
            pass
        return None

    def _is_positive_number(self, v):
        return isinstance(v, (int, float)) and v > 0

    def handle_result(self,info):
        if info:
            self.update_display(info)
            self.status_label.config(text="✅ Auto detection completed.",fg="#00ff00")
        else:
            self.status_label.config(text="❌ Auto-detection failed. Enter data manually.",fg="#ff4444")

    # -------------------- Display --------------------
    def update_display(self,info):
        # Normalize and ensure numeric types where possible
        self.current_data = info or {}
        try:
            self.save_data(self.current_data)
        except Exception:
            pass

        # safe format helper
        def safe_int(val, suffix=""):
            if isinstance(val, (int, float)):
                return f"{int(val):,}{suffix}"
            return "Unknown"
        def safe_percent(val):
            if isinstance(val, (int, float)):
                return f"{float(val):.1f}%"
            return "Unknown"

        status_map = {1:"Discharging",2:"AC Power",3:"Fully Charged",6:"Charging"}
        self.charge_label.config(text=safe_percent(info.get('percent')))
        self.state_label.config(text=status_map.get(info.get('status'),"Unknown"))
        self.power_label.config(text="Plugged In" if info.get('plugged') else "On Battery")
        self.design_label.config(text=safe_int(info.get('design_capacity')," mWh"))
        self.full_label.config(text=safe_int(info.get('full_capacity')," mWh"))
        self.cycle_label.config(text=safe_int(info.get('cycle_count')))
        # ensure health is numeric and capped
        health = info.get('health_percent')
        if isinstance(health, (int, float)):
            health = float(max(0.0, min(100.0, health)))
        self.health_label.config(text=safe_percent(health))
        self.draw_health_bar(health if isinstance(health, (int, float)) else 0.0)
        self.assess_label.config(**self.assessment(health if isinstance(health, (int, float)) else 0.0))
        self.update_label.config(text="Last updated: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def draw_health_bar(self,percent):
        self.health_canvas.delete("all")
        width=350;height=30
        p = percent if isinstance(percent, (int, float)) else 0.0
        p = max(0.0, min(100.0, float(p)))
        fill=max(5,int(width*(p/100.0)))
        colors=['#ff4444','#ffaa00','#ffff00','#aaff00','#00ff00']
        color=colors[min(4,max(0,int(p/25)))]
        self.health_canvas.create_rectangle(0,0,width,height,fill='#3a3a3a',outline='')
        self.health_canvas.create_rectangle(0,0,fill,height,fill=color,outline='')
        text_color='black' if p>50 else 'white'
        self.health_canvas.create_text(width/2,height/2,
                                       text=f"{p:.1f}% Health",
                                       fill=text_color,font=('Arial',10,'bold'))

    def assessment(self,percent):
        p = percent if isinstance(percent, (int, float)) else 0.0
        if p>=90: return {'text':"Assessment: Excellent ✅",'fg':'#00ff00'}
        if p>=80: return {'text':"Assessment: Very Good 👍",'fg':'#aaff00'}
        if p>=70: return {'text':"Assessment: Good 👌",'fg':'#ffff00'}
        if p>=60: return {'text':"Assessment: Fair ⚠️",'fg':'#ffaa00'}
        if p>=50: return {'text':"Assessment: Poor ❗",'fg':'#ff6600'}
        return {'text':"Assessment: Critical 🔴",'fg':'#ff4444'}

    # -------------------- Manual Entry --------------------
    def submit_manual(self):
        try:
            design = int(self.entries['design'].get() or 0)
            full = int(self.entries['full'].get() or 0)
            cycles = int(self.entries['cycle'].get() or 0)
            percent = int(self.entries['percent'].get() or 0)
            health = (full/design*100) if design and full else None
            if health is not None:
                health = float(max(0.0, min(100.0, health)))
            info={'design_capacity':design if design>0 else None,'full_capacity':full if full>0 else None,
                  'cycle_count':cycles if cycles>0 else None,'percent':percent if percent>=0 else None,
                  'status':1,'plugged':False,'health_percent':health}
            self.update_display(info)
            self.status_label.config(text="✅ Manual data entered.",fg="#00ff00")
        except ValueError:
            messagebox.showerror("Error","Please enter valid numbers in all fields.")

    # -------------------- Utilities --------------------
    def save_current_data(self):
        if self.current_data:
            self.save_data(self.current_data)
            messagebox.showinfo("Saved","Battery data saved successfully!")
        else:
            messagebox.showwarning("Warning","No data to save.")

    def clear_data(self):
        if messagebox.askyesno("Confirm","Clear all battery data?"):
            self.current_data=None
            if os.path.exists('battery_data.json'): os.remove('battery_data.json')
            for lbl in [self.charge_label,self.state_label,self.power_label,
                        self.design_label,self.full_label,self.cycle_label,
                        self.health_label,self.assess_label]:
                lbl.config(text="--")
            self.draw_health_bar(0)
            self.update_label.config(text="Last updated: --")
            self.status_label.config(text="Data cleared.",fg="#ffff00")

    def generate_report(self):
        def run():
            try:
                out_path = os.path.join(os.environ.get("USERPROFILE", "."), "battery-report.html")
                r = subprocess.run(["powercfg","/batteryreport","/output", out_path],
                                   capture_output=True,text=True,timeout=30)
                if r.returncode == 0 or os.path.exists(out_path):
                    messagebox.showinfo("Report",f"Report saved to:\n{out_path}")
                    try:
                        os.startfile(out_path)
                    except Exception:
                        pass
                else:
                    messagebox.showerror("Error","Could not generate report.")
            except Exception as e:
                messagebox.showerror("Error",str(e))
        threading.Thread(target=run,daemon=True).start()

def main():
    root = tk.Tk()
    BatteryHealthGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
