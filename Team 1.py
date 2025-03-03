# Install necessary libraries
%pip install pandas openpyxl nltk gensim scikit-learn matplotlib

# Import necessary libraries
import pandas as pd
import openpyxl
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import matplotlib.pyplot as plt

# Download NLTK data
nltk.download('punkt')
nltk.download('stopwords')

# Load the Excel file (ensure the path is correct)
file_path = r'C:\Users\abhin\Downloads\2024_Myeloma All Inquiries and Potental AE Reporting.xlsx' #Replace with file path according to user
df = pd.read_excel(file_path)

# Check if the necessary columns exist
if 'Community Oncologist Question' not in df.columns or 'Expert Response' not in df.columns:
    raise ValueError("Required columns not found in the DataFrame.")

# Separate texts from each column
community_questions = df['Community Oncologist Question'].dropna().tolist()
expert_responses = df['Expert Response'].dropna().tolist()

# Preprocess the text data
stop_words = set(stopwords.words('english'))
stop_words.add('please')  # Add custom stop word
stop_words.add('get') 
stop_words.add('input') 

def preprocess(text):
    tokens = word_tokenize(text.lower())
    tokens = [word for word in tokens if (word.isalpha() or '-' in word) and word not in stop_words]
    return tokens

# Preprocess texts from each column separately
processed_community_questions = [preprocess(text) for text in community_questions]
processed_expert_responses = [preprocess(text) for text in expert_responses]

# Function to filter sentences containing specified terms
def filter_sentences_with_terms(texts, terms):
    filtered_sentences = []
    seen_initial_strings = set()
    for tokens in texts:
        if any(term in tokens for term in terms):
            initial_string = ' '.join(tokens[:5])  # Compare the first 5 tokens as the initial string
            if initial_string not in seen_initial_strings:
                filtered_sentences.append(tokens)
                seen_initial_strings.add(initial_string)
    return filtered_sentences

# Define the sets of terms to filter
terms_sets = {
    'Daratumumab': ['daratumumab', 'darzalex', 'dara', 'iv dara', 'Daratumumab', 'Darzalex', 'Dara', 'IV Dara'],
    'Sarclisa': ['isatuximab', 'sarclisa', 'isa', 'iv isa', 'Isatuximab', 'Sarclisa', 'Isa', 'IV Isa'],
    'Revlimid': ['revlimid', 'lenalidomide', 'len', 'Revlimid', 'Lenalidomide', 'Len'],
    'Carvykti': ['carvykti', 'ciltacabtagene', 'car-t', 'cart', 'ciltacel', 'cilta-cel', 'bcma car-t', 'bcma cart', 'CAR-T', 'CART', 'Carvykti', 'Ciltacabtagene', 'Ciltacel', 'CAR T'],
    #'Abecma': ['abecma', 'idecabtagene', 'ide-cel','Abecma', 'Idecabtagene', 'IDE-Cel'],
    'Tecvayli': ['tecvayli', 'teclistamab', 'tec', 'bispecific', 'bcma bispecific', 'biabs', 'Tecvayli', 'Teclistamab', 'Tec', 'Bispecific', 'BCMA Bispecific', 'BiAbs'],
    'Talvey': ['talvey', 'talquetamab', 'tal', 'bispecific', 'GPRC5D bispecific', 'biAbs', 'Talvey', 'Talquetamab', 'Tal', 'Bispecific', 'GPRC5D Bispecific', 'BiAbs'],
    #'Elranatamab': ['elranatamab', 'elra','Elranatamab', 'Elra'],
    #'Cevostamab': ['cevostamab', 'cevo', 'Cevostamab', 'Cevo'],
}

# Create a dictionary to store the filtered sentences for each set of terms
filtered_sentences_dict = {}
term_occurrences = {}

# Filter the preprocessed texts for each set of terms
for key, terms in terms_sets.items():
    community_questions_with_terms = filter_sentences_with_terms(processed_community_questions, terms)
    expert_responses_with_terms = filter_sentences_with_terms(processed_expert_responses, terms)
    filtered_sentences_dict[key] = {
        "Community Oncologist Questions": community_questions_with_terms,
        "Expert Responses": expert_responses_with_terms
    }
    term_occurrences[key] = {
        "Community Oncologist Questions": len(community_questions_with_terms),
        "Expert Responses": len(expert_responses_with_terms)
    }

# Display the results for each set of terms
for key, categories in filtered_sentences_dict.items():
    print(f"\nSentences mentioning terms related to '{key}':")
    for category, sentences in categories.items():
        print(f"\n{category}:")
        for sentence in sentences:
            print(sentence)

# Display the term occurrences
print("\nTerm Occurrences:")
for key, counts in term_occurrences.items():
    print(f"\nTerms related to '{key}':")
    for category, count in counts.items():
        print(f"{category}: {count}")

# Create pie charts for term occurrences
community_labels = []
community_sizes = []
expert_labels = []
expert_sizes = []

for key, counts in term_occurrences.items():
    community_labels.append(key)
    community_sizes.append(counts["Community Oncologist Questions"])
    expert_labels.append(key)
    expert_sizes.append(counts["Expert Responses"])

# Pie chart for Community Oncologist Questions
plt.figure(figsize=(8, 8))
plt.pie(community_sizes, labels=community_labels, autopct='%1.1f%%', startangle=140, textprops={'fontsize': 14}, labeldistance=0.7)
plt.title("Share of Voice of Treatments - Community Oncologists", fontsize=16)
plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
plt.legend(loc="upper left", bbox_to_anchor=(1, 0.5))
plt.show()

# Pie chart for Expert Responses
plt.figure(figsize=(8, 8))
plt.pie(expert_sizes, labels=expert_labels, autopct='%1.1f%%', startangle=140, textprops={'fontsize': 14}, labeldistance=0.7)
plt.title("Share of Voice of Treatments - Expert Responses", fontsize=16)
plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
plt.legend(loc="upper left", bbox_to_anchor=(1, 0.5))
plt.show()
