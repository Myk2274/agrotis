
# main.py
from sap import connect_sap, sap_transaction
import sys


def main():
    print("Conectando ao SAP...")
    session = connect_sap()
    print("Conectado com sucesso!")

    print("Executando transação...")
    
    cliente = sys.argv[1]
    
    sap_transaction(session, int(cliente))
    

if __name__ == "__main__":
    main()
