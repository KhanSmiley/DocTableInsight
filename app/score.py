import sys
import pathlib
import csv
from pathlib import Path
backend_ai_path = pathlib.Path(__file__).parent.parent

print("parent_dir :",backend_ai_path)
# Append 'config' to the parent directory
sys.path.append(str(backend_ai_path))

file =  "path/cisco/queries_to_check/cisco_queries_r4.xlsx"
file_path = Path(backend_ai_path/file) #db_creation_inputs.get("chunk_folder")
print("File Path: ", file_path)

import pandas as pd
from rouge_score import rouge_scorer

# Function to calculate ROUGE-L score
def calculate_rouge_l(reference: str, hypothesis: str) -> float:
    scorer = rouge_scorer.RougeScorer(['rougeL'], use_stemmer=True)
    scores = scorer.score(reference, hypothesis)
    return scores['rougeL'].fmeasure

# Mock function to calculate RAGAS score
def calculate_ragas(reference: str, hypothesis: str) -> float:
    # Simplified placeholder logic for RAGAS calculation
    # You need to replace this with actual RAGAS implementation logic
    # This could involve semantic similarity, relevance scoring, etc.
    # Here we just return a dummy score
    return 0.85  # Dummy score for demonstration

# Read the Excel file
# file_path = './path/cisco/queries_to_check/cisco_queries_r4.xlsx'  # Replace with the actual file path if needed
df = pd.read_excel(file_path, sheet_name="FY24_Sales_Comp_FAQ")

# Assuming the columns are named as in the description
ground_truths = df['ground_truth']
responses = df['response']

# Calculate scores
results = []
for ground_truth, response in zip(ground_truths, responses):
    rouge_l = calculate_rouge_l(ground_truth, response)
    ragas = calculate_ragas(ground_truth, response)
    results.append((rouge_l, ragas))

# Add results to the DataFrame
df['ROUGE-L'] = [result[0] for result in results]
df['RAGAS'] = [result[1] for result in results]

# Save the DataFrame with scores to a new Excel file
output_file_path = 'C:/Users/Admin/OneDrive - Relanto/Desktop/Projects/AI_lab/backend-ai/path/cisco/queries_to_check/evaluated_results.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Results saved to {output_file_path}")