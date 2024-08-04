import tkinter as tk
import speech_recognition as sr
from commands import process_command_func
from tasks import speak, greet

class VoxBuddyInterface:
    def __init__(self, root, process_command_func):
        self.root = root
        self.root.title("Vox Buddy: Your Personal Assistant")

        self.dialog_text = tk.Text(root, height=40, width=50, wrap=tk.WORD)
        self.dialog_text.pack(side=tk.RIGHT)

        self.result_text = tk.Text(root, height=40, width=50, wrap=tk.WORD)
        self.result_text.pack(side=tk.LEFT)

        self.process_command_func = process_command_func

        self.recognizer = sr.Recognizer()

        greet_response = greet()
        self.result_text.insert(tk.END, f"Assistant: {greet_response}\n\n")
        self.root.update()
        speak(greet_response)

        self.listen_for_audio()

    def listen_for_audio(self):
        self.result_text.insert(tk.END, "Listening...\n\n")
        speak("listening.....")
        self.root.update()

        with sr.Microphone() as source:
            self.recognizer.adjust_for_ambient_noise(source)

            try:
                audio = self.recognizer.listen(source, timeout=10)
                query = self.recognizer.recognize_google(audio)
                self.dialog_text.insert(tk.END, f"You said : {query}\n\n")
                self.dialog_text.see(tk.END)
                self.root.update()  

                response = self.process_command_func(query)
                self.result_text.insert(tk.END, f"Voice Buddy : {response}\n\n")
                self.dialog_text.see(tk.END)
                self.root.update()  

                speak(response)

                print(f"You said : {query}")
                print(f"Voice Buddy : {response}")

            except sr.WaitTimeoutError:
                print("Listening timed out. Please try again.")
            except sr.UnknownValueError:
                print("No speech detected. Please try again.")
            except sr.RequestError as e:
                print(f"Error connecting to Google API: {e}")
            except Exception as e:
                print(f"The Error Occurred: {e}")

        self.listen_for_audio()

if __name__ == "__main__":
    root = tk.Tk()
    interface = VoxBuddyInterface(root, process_command_func)
    root.mainloop()
