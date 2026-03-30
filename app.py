import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from pdf_processor import PDFProcessor
from email_sender import EmailSender
import threading
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

# Set appearance and theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("보험료 연말정산 시스템")
        self.geometry("800x800")

        # Grid configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Tabview
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_pdf = self.tabview.add("PDF 분할 & 마스킹")
        self.tab_email = self.tabview.add("이메일 발송")

        self.setup_pdf_tab()
        self.setup_email_tab()

        # Log View (Shared)
        self.log_view = ctk.CTkTextbox(self, height=120)
        self.log_view.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.log_view.insert("0.0", "시스템 로그\n")
        self.log_view.configure(state="disabled")

    def setup_pdf_tab(self):
        self.tab_pdf.grid_columnconfigure(0, weight=1)

        # Title
        title = ctk.CTkLabel(self.tab_pdf, text="PDF 분할 & 마스킹 설정", font=ctk.CTkFont(size=18, weight="bold"))
        title.grid(row=0, column=0, padx=20, pady=10)

        # File Selection Frame
        file_frame = ctk.CTkFrame(self.tab_pdf)
        file_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(file_frame, text="입력 PDF:").grid(row=0, column=0, padx=10, pady=10)
        self.file_entry = ctk.CTkEntry(file_frame, placeholder_text="PDF 파일 선택...")
        self.file_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(file_frame, text="찾아보기", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)

        # Output Selection Frame
        output_frame = ctk.CTkFrame(self.tab_pdf)
        output_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        output_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(output_frame, text="저장 폴더:").grid(row=0, column=0, padx=10, pady=10)
        self.output_entry = ctk.CTkEntry(output_frame, placeholder_text="저장 경로 선택...")
        self.output_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(output_frame, text="찾아보기", command=self.browse_folder).grid(row=0, column=2, padx=10, pady=10)

        # Action Button
        self.process_button = ctk.CTkButton(self.tab_pdf, text="PDF 처리 시작", command=self.start_pdf_processing, height=40)
        self.process_button.grid(row=3, column=0, padx=20, pady=20, sticky="ew")

    def setup_email_tab(self):
        self.tab_email.grid_columnconfigure(0, weight=1)

        # SMTP Settings Frame
        smtp_frame = ctk.CTkFrame(self.tab_email)
        smtp_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        smtp_frame.grid_columnconfigure((1, 3), weight=1)

        ctk.CTkLabel(smtp_frame, text="SMTP 서버:").grid(row=0, column=0, padx=5, pady=5)
        self.smtp_server = ctk.CTkEntry(smtp_frame, placeholder_text="smtp.gmail.com")
        self.smtp_server.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_server.insert(0, "smtp.gmail.com")

        ctk.CTkLabel(smtp_frame, text="포트:").grid(row=0, column=2, padx=5, pady=5)
        self.smtp_port = ctk.CTkEntry(smtp_frame, placeholder_text="587", width=60)
        self.smtp_port.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        self.smtp_port.insert(0, "587")

        ctk.CTkLabel(smtp_frame, text="계정(Email):").grid(row=1, column=0, padx=5, pady=5)
        self.email_user = ctk.CTkEntry(smtp_frame, placeholder_text="example@gmail.com")
        self.email_user.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(smtp_frame, text="발송자 이름:").grid(row=1, column=2, padx=5, pady=5)
        self.sender_name = ctk.CTkEntry(smtp_frame, placeholder_text="해외건설협회 경영지원실 정효진")
        self.sender_name.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        self.sender_name.insert(0, "해외건설협회 경영지원실 정효진")

        ctk.CTkLabel(smtp_frame, text="비밀번호:").grid(row=2, column=0, padx=5, pady=5)
        self.email_pass = ctk.CTkEntry(smtp_frame, show="*", placeholder_text="앱 비밀번호")
        self.email_pass.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(smtp_frame, text="발송자 메일:").grid(row=2, column=2, padx=5, pady=5)
        self.display_email = ctk.CTkEntry(smtp_frame, placeholder_text="보이는 이메일 주소")
        self.display_email.grid(row=2, column=3, padx=5, pady=5, sticky="ew")
 
        # Help hint for Gmail App Password
        help_hint = ctk.CTkLabel(smtp_frame, text="*구글 계정은 '앱 비밀번호' 필수", font=ctk.CTkFont(size=10), text_color="gray")
        help_hint.grid(row=3, column=1, padx=5, pady=0, sticky="en")

        # Excel Mapping Frame
        mapping_frame = ctk.CTkFrame(self.tab_email)
        mapping_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        mapping_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(mapping_frame, text="명단(Excel):").grid(row=0, column=0, padx=10, pady=10)
        self.excel_entry = ctk.CTkEntry(mapping_frame, placeholder_text="연번, 성명, 이메일이 포함된 엑셀 선택...")
        self.excel_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(mapping_frame, text="찾아보기", command=self.browse_excel).grid(row=0, column=2, padx=10, pady=10)

        # Email Template Frame
        template_frame = ctk.CTkFrame(self.tab_email)
        template_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        template_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(template_frame, text="제목:").grid(row=0, column=0, padx=10, pady=5)
        self.email_subject = ctk.CTkEntry(template_frame)
        self.email_subject.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.email_subject.insert(0, "[경영지원실] {성명}님 보험료 연말정산 산출내역서입니다.")

        ctk.CTkLabel(template_frame, text="본문:").grid(row=1, column=0, padx=10, pady=5, sticky="n")
        self.email_body = ctk.CTkTextbox(template_frame, height=80)
        self.email_body.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.email_body.insert("0.0", "안녕하세요, 경영지원실입니다.\n{성명}님의 보험료 연말정산 산출내역서를 보내드립니다.\n첨부파일을 확인해 주세요.\n\n감사합니다.")

        # Action Buttons
        button_frame = ctk.CTkFrame(self.tab_email, fg_color="transparent")
        button_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        button_frame.grid_columnconfigure((0, 1), weight=1)

        self.test_email_button = ctk.CTkButton(button_frame, text="테스트 메일 전송", command=self.send_test_email, fg_color="gray")
        self.test_email_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.send_all_button = ctk.CTkButton(button_frame, text="파일 매칭 및 전체 발송", command=self.start_email_sending)
        self.send_all_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    # Functions
    def log(self, message):
        self.log_view.configure(state="normal")
        self.log_view.insert(tk.END, message + "\n")
        self.log_view.see(tk.END)
        self.log_view.configure(state="disabled")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)
            if not self.output_entry.get():
                self.output_entry.insert(0, os.path.join(os.path.dirname(filename), "output"))

    def browse_folder(self):
        foldername = filedialog.askdirectory()
        if foldername:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, foldername)

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, filename)

    def start_pdf_processing(self):
        input_path = self.file_entry.get()
        output_dir = self.output_entry.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "PDF 파일을 선택하세요.")
            return
        self.process_button.configure(state="disabled", text="처리 중...")
        threading.Thread(target=self.run_pdf_task, args=(input_path, output_dir), daemon=True).start()

    def run_pdf_task(self, input_path, output_dir):
        try:
            processor = PDFProcessor(input_path, output_dir, log_callback=self.log)
            processor.process()
            messagebox.showinfo("완료", "PDF 분할 및 마스킹이 완료되었습니다.")
        except Exception as e:
            self.log(f"오류: {str(e)}")
        finally:
            self.process_button.configure(state="normal", text="PDF 처리 시작")

    def send_test_email(self):
        user = self.email_user.get()
        if not user:
            messagebox.showerror("Error", "본인의 이메일 주소를 입력하세요.")
            return
        self.log(f"테스트 메일을 {user}로 전송 시도 중...")
        threading.Thread(target=self.run_send_test, daemon=True).start()

    def run_send_test(self):
        sender = EmailSender(self.smtp_server.get(), int(self.smtp_port.get()), self.email_user.get(), self.email_pass.get())
        success, msg = sender.send_email(
            self.email_user.get(), 
            "테스트 메일입니다.", 
            "시스템 이메일 발송 기능이 정상 작동합니다.",
            sender_name=self.sender_name.get(),
            display_email=self.display_email.get()
        )
        if success:
            self.log("테스트 메일 발송 성공!")
            messagebox.showinfo("성공", "테스트 메일이 발송되었습니다.")
        else:
            self.log(f"테스트 메일 발송 실패: {msg}")
            messagebox.showerror("실패", f"발송 실패: {msg}")

    def start_email_sending(self):
        excel_path = self.excel_entry.get()
        output_dir = self.output_entry.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "엑셀 명단 파일을 선택하세요.")
            return
        if not output_dir or not os.path.exists(output_dir):
            messagebox.showerror("Error", "PDF 결과물이 저장된 폴더를 확인하세요.")
            return
        self.send_all_button.configure(state="disabled", text="발송 중...")
        threading.Thread(target=self.run_email_batch, args=(excel_path, output_dir), daemon=True).start()

    def run_email_batch(self, excel_path, output_dir):
        try:
            self.log("엑셀 파일을 불러오는 중...")
            df = pd.read_excel(excel_path)
            
            files = [f for f in os.listdir(output_dir) if f.endswith('.pdf')]
            self.log(f"폴더 내 총 {len(files)}건의 발송 준비를 시작합니다.")
            
            # Prepare email data
            email_tasks = []
            for filename in files:
                try:
                    basename = os.path.splitext(filename)[0]
                    parts = basename.split('-')
                    if len(parts) < 2: continue
                    serial = parts[0].strip()
                    name = parts[1].strip()
                    
                    match = df[(df.iloc[:,0].astype(str) == serial) & (df.iloc[:,1].astype(str) == name)]
                    
                    if not match.empty:
                        target_email = str(match.iloc[0, 2]).strip()
                        subject = self.email_subject.get().replace("{성명}", name)
                        body = self.email_body.get("0.0", tk.END).replace("{성명}", name)
                        pdf_path = os.path.join(output_dir, filename)
                        
                        email_tasks.append({
                            'receiver': target_email,
                            'subject': subject,
                            'body': body,
                            'attachment': pdf_path,
                            'name': name
                        })
                except:
                    pass

            if not email_tasks:
                self.log("발송할 대상자를 찾지 못했습니다.")
                return

            self.log(f"매칭 완료! 병렬 발송(10개 스레드)을 시작합니다...")
            
            # Email sending worker function
            def send_worker(task):
                # Each worker needs its own connection for parallel sending
                w_sender = EmailSender(
                    self.smtp_server.get(), 
                    int(self.smtp_port.get()), 
                    self.email_user.get(), 
                    self.email_pass.get()
                )
                success, error_msg = w_sender.send_email(
                    task['receiver'], 
                    task['subject'], 
                    task['body'], 
                    task['attachment'],
                    sender_name=self.sender_name.get(),
                    display_email=self.display_email.get()
                )
                return success, task['name'], error_msg

            success_count = 0
            total_tasks = len(email_tasks)
            # Using max_workers=10 for maximum speed
            with ThreadPoolExecutor(max_workers=10) as executor:
                # Use as_completed for real-time logging as each thread finishes
                future_to_name = {executor.submit(send_worker, task): task['name'] for task in email_tasks}
                
                for i, future in enumerate(as_completed(future_to_name), 1):
                    name = future_to_name[future]
                    try:
                        success, name, error_msg = future.result()
                        if success:
                            success_count += 1
                            self.log(f"[{i}/{total_tasks}] 성공: {name}")
                        else:
                            self.log(f"[{i}/{total_tasks}] 실패: {name} ({error_msg})")
                    except Exception as exc:
                        self.log(f"[{i}/{total_tasks}] 오류: {name} ({exc})")

            self.log(f"전체 발송 완료! (성공: {success_count}/{total_tasks})")
            messagebox.showinfo("완료", f"전체 발송이 완료되었습니다. (성공: {success_count})")
            
        except Exception as e:
            self.log(f"오류: {str(e)}")
            messagebox.showerror("Error", f"작업 중 오류 발생: {str(e)}")
        finally:
            self.send_all_button.configure(state="normal", text="파일 매칭 및 전체 발송")

if __name__ == "__main__":
    app = App()
    app.mainloop()
