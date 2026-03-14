from deep_translator import GoogleTranslator

try:
    word = "apple"
    translated = GoogleTranslator(source='en', target='zh-TW').translate(word)
    print(f"Word: {word}, Translated: {translated}")
    
    word2 = "trajectory"
    translated2 = GoogleTranslator(source='en', target='zh-TW').translate(word2)
    print(f"Word: {word2}, Translated: {translated2}")
except Exception as e:
    print(f"Error: {e}")
