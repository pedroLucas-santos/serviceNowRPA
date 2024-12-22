from flask import Flask, jsonify
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from multiprocessing import Process, Pipe

app = Flask(__name__)

def tkinter_file_dialog(conn):
    """Função para abrir o tkinter em um processo separado."""
    root = Tk()
    root.withdraw()
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    conn.send(file_path)
    conn.close()

@app.route('/select-file', methods=['GET'])
def select_file():
    try:
        # Cria um Pipe para comunicação entre processos
        parent_conn, child_conn = Pipe()
        # Inicia o processo com tkinter
        p = Process(target=tkinter_file_dialog, args=(child_conn,))
        p.start()
        p.join()  # Aguarda o processo terminar
        file_path = parent_conn.recv()  # Recebe o resultado
        
        if file_path:
            return jsonify({'file_path': file_path})
        else:
            return jsonify({'error': 'No file selected'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
