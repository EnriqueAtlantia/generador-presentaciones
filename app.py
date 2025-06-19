# app.py
from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import openai
import json
import io

app = Flask(__name__)
CORS(app)

@app.route('/api/generar', methods=['POST'])
def generar_presentacion():
    try:
        api_key = request.form.get('api_key')
        excel_file = request.files.get('excel_file')
        storytelling_file = request.files.get('storytelling_file')
        structure_file = request.files.get('structure_file')

        if not all([api_key, excel_file, storytelling_file, structure_file]):
            return jsonify({'error': 'Faltan datos o archivos'}), 400

        client = openai.OpenAI(api_key=api_key)

        df = pd.read_excel(excel_file.stream)
        storytelling_prompt = storytelling_file.read().decode('utf-8')
        ppt_structure_prompt = structure_file.read().decode('utf-8')

        data_for_prompt = df.head(50).to_markdown(index=False)
        system_prompt = ppt_structure_prompt
        user_prompt = f"{storytelling_prompt}\n\n**DATOS CUANTITATIVOS A ANALIZAR:**\n{data_for_prompt}"

        response = client.chat.completions.create(
            model="gpt-4o",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ]
        )
        presentation_plan = json.loads(response.choices[0].message.content)

        markdown_content = []
        main_title = presentation_plan.get("titulo_presentacion", "Título")
        markdown_content.append(f"# {main_title}\n\n---\n")

        for slide in presentation_plan.get("diapositivas", []):
            markdown_content.append(f"## {slide.get('titulo', 'Diapositiva sin título')}")
            for point in slide.get("puntos_clave", []):
                markdown_content.append(f"* {point}")
            if slide.get("grafico"):
                markdown_content.append(f"\n_[Nota: Insertar gráfico: {slide['grafico'].get('titulo_grafico', '')}]_")
            markdown_content.append("\n---\n")

        final_markdown_text = "\n".join(markdown_content)

        return jsonify({'markdown': final_markdown_text})

    except Exception as e:
        return jsonify({'error': str(e)}), 500