import pandas as pd
import genanki
import requests
import os
import re
import warnings
from deep_translator import GoogleTranslator

# Suppress openpyxl warning: "Workbook contains no default style"
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

class AnkiConverter:
    def __init__(self, excel_path, output_path, deck_name, include_tts=False, 
                 show_translation=True, show_pos=True, show_explanation=True):
        self.excel_path = excel_path
        self.output_path = output_path
        self.deck_name = deck_name
        self.include_tts = include_tts
        self.show_translation = show_translation
        self.show_pos = show_pos
        self.show_explanation = show_explanation
        
        self.translator = GoogleTranslator(source='en', target='zh-TW')
        self.model_id = 1607392319
        self.deck_id = 2059400110
        
        # Build dynamic template based on user choices
        explanation_part = '<div class="explanation">{{#Explanation}}{{Explanation}}{{/Explanation}} {{#POS}}(<i>{{POS}}</i>){{/POS}}</div>' if (show_explanation or show_pos) else ''
        translation_part = '<div class="translation">{{#Translation}}{{Translation}}{{/Translation}}</div>' if show_translation else ''
        
        self.model = genanki.Model(
            self.model_id,
            'Cambridge Dictionary Model v0.0.3',
            fields=[
                {'name': 'Word'},
                {'name': 'POS'},
                {'name': 'Translation'},
                {'name': 'Explanation'},
                {'name': 'TTS'},
            ],
            templates=[
                {
                    'name': 'Card 1',
                    'qfmt': f'''
                        <div class="card-content">
                            {explanation_part}
                            {translation_part}
                            <div class="type-box">{{{{type:Word}}}}</div>
                            <div class="tts">{{{{TTS}}}}</div>
                        </div>
                    ''',
                    'afmt': f'''
                        <div class="card-content">
                            {explanation_part}
                            {translation_part}
                            <hr id="answer">
                            <div class="word">{{{{Word}}}}</div>
                            <div class="type-box">{{{{type:Word}}}}</div>
                        </div>
                    ''',
                },
            ],
            css='''
                .card {
                    font-family: "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                    background-color: #f1f5f9;
                    margin: 0;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    min-height: 100vh;
                    padding: 20px;
                }
                .card-content {
                    background: white;
                    padding: 40px;
                    border-radius: 20px;
                    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
                    width: auto;
                    min-width: 400px;
                    max-width: 90vw;
                    text-align: center;
                    overflow: visible;
                }
                .explanation {
                    color: #64748b;
                    font-size: 18px;
                    line-height: 1.5;
                    margin-bottom: 15px;
                }
                .translation {
                    color: #0d9488;
                    font-size: 28px;
                    font-weight: 700;
                    margin-bottom: 25px;
                }
                .word {
                    color: #1e293b;
                    font-size: 36px;
                    font-weight: 800;
                    margin: 20px 0;
                }
                .type-box {
                    margin-top: 20px;
                    overflow: visible;
                    display: inline-block;
                    width: 100%;
                }
                .tts {
                    margin-top: 25px;
                }
                #answer {
                    border: none;
                    border-top: 2px solid #e2e8f0;
                    margin: 30px 0;
                }
                
                /* Anki type-in box and feedback table styling */
                #typeans {
                    width: 100% !important;
                    box-sizing: border-box;
                    padding: 12px;
                    font-size: 20px;
                    border: 2px solid #cbd5e1;
                    border-radius: 12px;
                    text-align: center;
                    outline: none;
                    display: inline-block;
                }
                
                /* Anki Feedback Table (the comparison) */
                table.typeans {
                    margin: 0 auto;
                    border-collapse: separate;
                    border-spacing: 0 4px;
                }
                
                #typeans code { font-family: "Consolas", monospace; font-size: 1.1em; }
                .typeGood { background-color: #dcfce7 !important; color: #166534 !important; padding: 4px 8px; border-radius: 4px; }
                .typeBad { background-color: #fee2e2 !important; color: #991b1b !important; padding: 4px 8px; border-radius: 4px; }
                .typeMissed { background-color: #fef9c3 !important; color: #854d0e !important; padding: 4px 8px; border-radius: 4px; }
                
                /* TTS Play Button Beautification */
                .replay-button svg {
                    width: 45px;
                    height: 45px;
                }
                .replay-button svg circle {
                    fill: #0d9488;
                    transition: fill 0.2s, transform 0.2s;
                }
                .replay-button svg path {
                    fill: white;
                }
                .replay-button:hover svg circle {
                    fill: #0f766e;
                    transform: scale(1.05);
                }
            '''
        )

    def fetch_word_data(self, word):
        """Fetch POS and English explanation from Free Dictionary API."""
        try:
            response = requests.get(f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}", timeout=5)
            if response.status_code == 200:
                data = response.json()[0]
                meanings = data.get('meanings', [])
                if meanings:
                    pos = meanings[0].get('partOfSpeech', '')
                    explanation = meanings[0].get('definitions', [{}])[0].get('definition', '')
                    return pos, explanation
        except Exception:
            pass
        return "", ""

    def translate_to_chinese(self, word):
        """Translate word to Traditional Chinese using deep-translator."""
        try:
            translated = self.translator.translate(word)
            return translated if translated else ""
        except Exception as e:
            print(f"Translation error: {e}")
            return ""

    def process(self, progress_callback=None):
        # Based on detailed inspection:
        # Row 0: Banner (Skip)
        # Row 1: Header (Skip)
        # Data starts from Row 2 (Index 2 in raw)
        
        # Load without headers to control indexing manually
        df = pd.read_excel(self.excel_path, header=None, skiprows=2)
        
        # Column 0: Word
        # Column 1: POS (noun/verb)
        # Column 2: POS (backup/internal)
        # Column 3: Translation (Traditional Chinese)
        # Column 4: Definition/Explanation (English)
        
        deck = genanki.Deck(self.deck_id, self.deck_name)
        
        total_rows = len(df)
        for i, row in df.iterrows():
            if len(row) < 1: continue
            
            raw_word = row.iloc[0]
            if pd.isna(raw_word): continue
            word = str(raw_word).strip()
            if not word: continue
            
            # Translation is at Index 3
            raw_trans = row.iloc[3] if len(row) > 3 else None
            translation = str(raw_trans).strip() if not pd.isna(raw_trans) and str(raw_trans).lower() != 'nan' else ""
            
            # POS is at Index 1
            raw_pos = row.iloc[1] if len(row) > 1 else None
            excel_pos = str(raw_pos).strip() if not pd.isna(raw_pos) and str(raw_pos).lower() != 'nan' else ""
            
            # Explanation is at Index 4
            raw_exp = row.iloc[4] if len(row) > 4 else None
            excel_exp = str(raw_exp).strip() if not pd.isna(raw_exp) and str(raw_exp).lower() != 'nan' else ""
            
            # Fetch missing info from APIs if Excel is empty
            api_pos, api_exp = self.fetch_word_data(word)
            
            # Use Excel data if exists, otherwise fallback to API
            final_translation = translation if translation else self.translate_to_chinese(word)
            final_pos = excel_pos if excel_pos else api_pos
            final_explanation = excel_exp if excel_exp else api_exp
            
            # Respect UI toggles
            field_pos = final_pos if self.show_pos else ""
            field_trans = final_translation if self.show_translation else ""
            field_exp = final_explanation if self.show_explanation else ""
            
            tts_tag = f"[anki:tts lang=en_US]{word}[/anki:tts]" if self.include_tts else ""
            
            note = genanki.Note(
                model=self.model,
                fields=[word, field_pos, field_trans, field_exp, tts_tag]
            )
            deck.add_note(note)
            
            if progress_callback:
                progress_callback(int((i + 1) / total_rows * 100))
        
        genanki.Package(deck).write_to_file(self.output_path)
        return True
