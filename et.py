import pandas as pd
from pandas import DataFrame
# from googletrans import Translator
import openai 

# temp
openai.api_key = 'YOUR_API_KEY_HERE'

def ask_gpt(user_input, model_name="gpt-4-turbo"):
    """
    Sends a message to the GPT model and returns the model's response.
    """
    completion = openai.chat.completions.create(
        model=model_name,
        messages=[
            {
                "role": "user",
                "content": user_input,
            },
        ],
    )
    return completion.choices[0].message.content

def translate_excel(input_path, output_path, src_lang='auto', dest_lang='en'):
    # Load the Excel file
    df = pd.read_excel(input_path)
    
    # Function to translate text
    def translate_text(text, dest_lang):
        try:
            if len(str(text)) > 0:
                #TODO: use re to filter out non-text cells
                
                prompt = f'Translate the text to {dest_lang}, only give me the translation. (or return the text as-is if no translation needed): \n' + text
                translated_text = ask_gpt(prompt)

                # print(prompt)
                print(text, '\n', translated_text, '\n')
                
                return translated_text
            return text
        except Exception as e:
            print(f"Error translating text: {text}. Error: {e}")
            return text

    # Apply the translation to each cell in the dataframe
    df_translated = df.map(func=lambda x: translate_text(str(x), dest_lang) if isinstance(x, str) else x)

    # Save the translated dataframe to a new Excel file
    df_translated.to_excel(output_path, index=False)
    print(f"Translated Excel file saved to {output_path}")


def run_script():
    # Translate entire sheet
    input_excel = 'example_sheets/qa_testing_table.xlsx'
    output_name = 'example_sheets/outputs/cn_test_0.xlsx'
    chinese = '简体中文'
    english = 'English'

    translate_excel(input_excel, output_name, src_lang='en', dest_lang=chinese)

if __name__=="__main__":
    run_script()