from flask import Flask, render_template, request, jsonify
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO
import base64
import traceback
import matplotlib
import yaml
matplotlib.use('Agg')  # Set the backend to Agg to avoid threading issues

app = Flask(__name__)

# Read in the Questions
# Path to the YAML file
file_path = 'ngs2020_questions.yaml'

try:
    # Open and load the YAML file
    with open(file_path, 'r') as file:
        TABLE_DESCRIPTIONS = yaml.safe_load(file)
        del TABLE_DESCRIPTIONS['PUMFID']
           
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found.")
except yaml.YAMLError as e:
    print(f"Error parsing YAML file: {e}")


@app.route('/check_tables')
def check_tables():
    try:
        with pd.ExcelFile("NGS_Tables.xlsx") as excel:
            existing_sheets = set(excel.sheet_names)
            available = [t for t in TABLE_DESCRIPTIONS if t in existing_sheets]
            missing = [t for t in TABLE_DESCRIPTIONS if t not in existing_sheets]
            return jsonify({
                'available': available,
                'missing': missing,
                'total_available': len(available),
                'total_missing': len(missing)
            })
    except Exception as e:
        return jsonify({'error': str(e)})

def _get_NGS_table(table_name):
    """Read a specific table from the Excel file"""
    try:
        df = pd.read_excel("NGS_Tables.xlsx", sheet_name=table_name)
        return df
    except Exception as e:
        print(f"Error reading table {table_name}: {e}")
        return pd.DataFrame()

def get_available_tables():
    """Get tables that actually exist in the Excel file"""
    try:
        with pd.ExcelFile("NGS_Tables.xlsx") as excel:
            existing_sheets = set(excel.sheet_names)
            available = [t for t in TABLE_DESCRIPTIONS if t in existing_sheets]
            missing = [t for t in TABLE_DESCRIPTIONS if t not in existing_sheets]
            if missing:
                print(f"Missing sheets: {missing}")
            return available
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def create_visualizations(df, table_name):
    visualizations = {}
    if df.empty:
        return visualizations
    
    try:
        # Verify required columns exist
        required_cols = ['Answer Categories', 'Weighted Frequency']
        if not all(col in df.columns for col in required_cols):
            print(f"Missing required columns in {table_name}")
            return visualizations

        # Clean and prepare data
        plot_df = df[
            (df['Answer Categories'].notna()) & 
            (df['Answer Categories'] != 'Total')
        ].copy()
        
        plot_df['Weighted Frequency'] = pd.to_numeric(
            plot_df['Weighted Frequency'].astype(str).str.replace(',', ''),
            errors='coerce'
        )
        plot_df = plot_df.dropna(subset=['Weighted Frequency'])
        
        if plot_df.empty:
            return visualizations

        # Convert to lists for JSON serialization
        categories = plot_df['Answer Categories'].tolist()
        frequencies = plot_df['Weighted Frequency'].tolist()

        # Create Pie Chart
        visualizations['pie_chart'] = {
            'data': [{
                'values': frequencies,
                'labels': categories,
                'type': 'pie',
                'textinfo': 'percent',
                'hoverinfo': 'label+percent+value',
                'hole': 0.3 if len(plot_df) > 5 else 0,
            }],
            'layout': {
                'title': f"{table_name}: Weighted Distribution",
                'margin': {'l': 20, 'r': 20, 'b': 20, 't': 40},
                'showlegend': True
            }
        }
        
        # Create Box Plot (only if we have numeric data)
        if pd.api.types.is_numeric_dtype(plot_df['Weighted Frequency']):
            visualizations['box_plot'] = {
                'data': [{
                    'y': frequencies,
                    'type': 'box',
                    'name': 'Weighted Frequency',
                    'boxpoints': 'all',
                    'jitter': 0.3,
                    'pointpos': -1.8
                }],
                'layout': {
                    'title': f"{table_name}: Weighted Frequency Distribution",
                    'yaxis': {'title': 'Weighted Frequency'},
                    'margin': {'l': 60, 'r': 20, 'b': 40, 't': 40}
                }
            }
        
        # Create Bar Chart
        visualizations['bar_chart'] = {
            'data': [{
                'x': categories,
                'y': frequencies,
                'type': 'bar',
                'marker': {'color': 'rgba(55, 128, 191, 0.7)'}
            }],
            'layout': {
                'title': f"{table_name}: Weighted Frequency by Category",
                'xaxis': {'title': 'Answer Categories'},
                'yaxis': {'title': 'Weighted Frequency'},
                'margin': {'l': 60, 'r': 20, 'b': 100, 't': 40}
            }
        }
        
    except Exception as e:
        print(f"Error in visualization: {str(e)}")
        traceback.print_exc()
        
    return visualizations

def get_basic_statistics(df):
    stats = {}
    if df.empty:
        return stats
    
    try:
        # Exclude 'code' column and only consider numeric columns
        numeric_cols = [col for col in df.columns 
                       if pd.api.types.is_numeric_dtype(df[col]) 
                       and col.lower() != 'code']
        
        for col in numeric_cols:
            stats[col] = {
                'mean': round(df[col].mean(), 2),
                'median': round(df[col].median(), 2),
                'min': round(df[col].min(), 2),
                'max': round(df[col].max(), 2),
                'std': round(df[col].std(), 2),
                'count': int(df[col].count())
            }
        return stats
    except Exception as e:
        print(f"Statistics error: {e}")
        return {}

@app.route('/')
def index():
    available_tables = get_available_tables()
    return render_template('index.html',
                         tables=available_tables,
                         table_descriptions=TABLE_DESCRIPTIONS,
                         total_tables=len(TABLE_DESCRIPTIONS),
                         available_count=len(available_tables))

@app.route('/get_table_data', methods=['POST'])
def get_table_data():
    try:
        table_name = request.json.get('table_name')
        if not table_name or table_name not in TABLE_DESCRIPTIONS:
            return jsonify({'success': False, 'error': 'Invalid table name'})
        
        df = _get_NGS_table(table_name)

        if df.empty:
            return jsonify({'success': False, 'error': 'Table is empty'})
        
        # Clean data - preserve code columns as strings
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str)
                # Only convert to numeric if not a code column
                if col.lower() != 'code':
                    df[col] = df[col].str.replace(',', '').str.replace('%', '')
                    try:
                        df[col] = pd.to_numeric(df[col])
                    except:
                        pass

        # Handle Total row - make Code empty if Answer Categories is 'Total'
        if 'Answer Categories' in df.columns and 'Code' in df.columns:
            total_mask = df['Answer Categories'].str.strip().str.lower() == 'total'
            df.loc[total_mask, 'Code'] = ''

        return jsonify({
            'success': True,
            'table_html': df.to_html(
                classes='table table-striped table-bordered table-hover',
                index=False,
                float_format=lambda x: '{:,.0f}'.format(x) if isinstance(x, (int, float)) and x.is_integer() else '{:,.2f}'.format(x) if isinstance(x, (int, float)) else str(x),
                na_rep=''  # This will make NaN values appear as empty
            ),
            'visualizations': create_visualizations(df, table_name),
            'statistics': get_basic_statistics(df),
            'table_description': TABLE_DESCRIPTIONS.get(table_name, '')
        })
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True, port=5000)