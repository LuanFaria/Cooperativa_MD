import os
import sys
import streamlit.web.cli as stcli

if __name__ == "__main__":
    # Descobre a pasta onde o executável está rodando
    base_path = os.path.dirname(__file__)
    app_path = os.path.join(base_path, "app_teste.py")
    
    # Simula o comando "streamlit run app.py" nativamente
    sys.argv = ["streamlit", "run", app_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())