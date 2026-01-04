import pandas as pd
import numpy as np
import win32com.client


class quiz:
    def __init__(self, csv_file):
        self.questions = pd.read_csv(csv_file)
        self.score = 0
        self.total = len(self.questions)

        # Windows Speech Engine
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")

    def speak(self, text):
        self.speaker.Speak(text)

    def start_quiz(self):
        print("Welcome to the quiz")
        self.speak("Welcome to the quiz")

        for index, row in self.questions.iterrows():
            print("\nQuestion:", row["question"])
            print("(A)", row["optionA"])
            print("(B)", row["optionB"])
            print("(C)", row["optionC"])
            print("(D)", row["optionD"])

            self.speak(row["question"])

            x = input("A/B/C/D: ").upper()

            if x == row["answer"]:
                print("Correct")
                self.speak("Correct")
                self.score += 1
            else:
                print("Incorrect")
                self.speak("Incorrect")

    def show_result(self):
        percentage = np.round((self.score / self.total) * 100, 2)

        print("\nQuiz completed")
        self.speak("Quiz completed")

        print("Your score is", self.score)
        self.speak(f"Your score is {self.score}")

        print("Your percentage is", percentage)
        self.speak(f"Your percentage is {percentage}")

        if percentage == 100:
            self.speak("Full marks")
            print("Full marks")
        elif percentage >= 80:
            self.speak("Grade A")
            print("Grade A")
        elif percentage >= 60:
            self.speak("Grade B")
            print("Grade B")
        else:
            self.speak("Grade C")
            print("Grade C")


quiz_1 = quiz(r"C:\Users\archit\Downloads\quiz_questions_10.csv")
quiz_1.start_quiz()
quiz_1.show_result()
