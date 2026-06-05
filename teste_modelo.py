"""
Script para testar o modelo de Rede Neural treinado
Carrega os pesos salvos e realiza predições na base de teste
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import json
from sklearn.metrics import confusion_matrix, classification_report
import seaborn as sns


class NnModel:
    """Modelo de Rede Neural para classificação"""
    
    def __init__(self, input_neurons, hidden_neurons, output_neurons):
        self.input_neurons = input_neurons
        self.hidden_neurons = hidden_neurons
        self.output_neurons = output_neurons
        
        # Inicializar pesos e biases
        self.w1 = np.zeros((input_neurons, hidden_neurons))
        self.b1 = np.zeros((1, hidden_neurons))
        self.w2 = np.zeros((hidden_neurons, output_neurons))
        self.b2 = np.zeros((1, output_neurons))
    
    def feedfoward(self, x):
        """Forward propagation"""
        z1 = x.dot(self.w1) + self.b1
        f1 = np.tanh(z1)
        z2 = f1.dot(self.w2) + self.b2
        output = np.tanh(z2)
        return output
    
    def carregar_pesos(self, caminho_pesos):
        """Carrega os pesos salvos do modelo"""
        print(f"Carregando pesos de: {caminho_pesos}")
        pesos_df = pd.read_csv(caminho_pesos)
        
        # Separar por camada
        w1_data = pesos_df[pesos_df['camada'] == 'w1'].drop('camada', axis=1).values
        b1_data = pesos_df[pesos_df['camada'] == 'b1'].drop('camada', axis=1).values
        w2_data = pesos_df[pesos_df['camada'] == 'w2'].drop('camada', axis=1).values
        b2_data = pesos_df[pesos_df['camada'] == 'b2'].drop('camada', axis=1).values
        
        self.w1 = w1_data
        self.b1 = b1_data
        self.w2 = w2_data
        self.b2 = b2_data
        
        print("Pesos carregados com sucesso!")
        print(f"W1: {self.w1.shape}, B1: {self.b1.shape}")
        print(f"W2: {self.w2.shape}, B2: {self.b2.shape}")


def main():
    """Função principal para testar o modelo"""
    
    print("=" * 70)
    print("TESTE DO MODELO DE REDE NEURAL TREINADO")
    print("=" * 70)
    
    # ========== 1. CARREGAR DADOS ==========
    print("\n1. Carregando dados...")
    try:
        x = np.load("/CARACTERES COMPLETO/X.npy")
        x = x.reshape(x.shape[0], -1)
        Y = np.load("/CARACTERES COMPLETO/Y_classe.npy")
        y_indices = np.argmax(Y, axis=1)
        print(f"   ✓ Dados carregados: {x.shape}")
    except FileNotFoundError as e:
        print(f"   ✗ Erro ao carregar dados: {e}")
        return
    
    # ========== 2. CARREGAR HIPERPARÂMETROS ==========
    print("\n2. Carregando hiperparâmetros...")
    try:
        with open("hiperparametros_finais.json", "r") as f:
            hyperparams = json.load(f)
        print("   ✓ Hiperparâmetros:")
        for key, value in hyperparams.items():
            print(f"     - {key}: {value}")
    except FileNotFoundError:
        print("   ✗ Arquivo hiperparametros_finais.json não encontrado")
        return
    
    # ========== 3. PREPARAR DADOS DE TESTE ==========
    print("\n3. Preparando dados de teste...")
    np.random.seed(42)
    n_samples = len(x)
    indices = np.arange(n_samples)
    np.random.shuffle(indices)
    
    train_end = int(0.7 * n_samples)
    val_end = int(0.8 * n_samples)
    test_indices = indices[val_end:]
    
    x_test = x[test_indices]
    y_test = y_indices[test_indices]
    
    print(f"   ✓ Dados de teste: {x_test.shape}")
    print(f"   ✓ Rótulos de teste: {y_test.shape}")
    
    # ========== 4. INICIALIZAR E CARREGAR MODELO ==========
    print("\n4. Inicializando modelo...")
    model = NnModel(
        input_neurons=hyperparams["input_neurons"],
        hidden_neurons=hyperparams["hidden_neurons"],
        output_neurons=hyperparams["output_neurons"]
    )
    
    try:
        model.carregar_pesos("pesos_finais.csv")
    except FileNotFoundError:
        print("   ✗ Arquivo pesos_finais.csv não encontrado")
        return
    
    # ========== 5. FAZER PREDIÇÕES ==========
    print("\n5. Realizando predições...")
    predictions = np.argmax(model.feedfoward(x_test), axis=1)
    accuracy = (predictions == y_test).mean()
    
    print(f"   ✓ Acurácia no Teste: {accuracy:.4f}")
    print(f"   ✓ Acertos: {(predictions == y_test).sum()} / {len(y_test)}")
    
    # ========== 6. RELATÓRIO DE CLASSIFICAÇÃO ==========
    print("\n6. Relatório de Classificação:")
    print("=" * 70)
    print(classification_report(
        y_test, 
        predictions, 
        target_names=[chr(65+i) for i in range(26)]
    ))
    
    # ========== 7. MATRIZ DE CONFUSÃO ==========
    print("\n7. Gerando Matriz de Confusão...")
    cm = confusion_matrix(y_test, predictions)
    
    fig, ax = plt.subplots(figsize=(14, 12))
    sns.heatmap(cm, annot=False, cmap='Blues', cbar=True, ax=ax)
    ax.set_xlabel('Predito')
    ax.set_ylabel('Real')
    ax.set_title('Matriz de Confusão - Teste do Modelo')
    ax.set_xticklabels([chr(65+i) for i in range(26)])
    ax.set_yticklabels([chr(65+i) for i in range(26)])
    plt.tight_layout()
    plt.savefig('matriz_confusao_teste.png', dpi=150)
    print("   ✓ Salvo em: matriz_confusao_teste.png")
    plt.show()
    
    # ========== 8. SALVAR RESULTADOS ==========
    print("\n8. Salvando resultados...")
    resultado_df = pd.DataFrame({
        'classe_real': y_test,
        'classe_prevista': predictions,
        'letra_real': [chr(65 + y) for y in y_test],
        'letra_predita': [chr(65 + p) for p in predictions],
        'acerto': predictions == y_test
    })
    
    resultado_df.to_csv('resultado_teste_final.csv', index=False)
    print(f"   ✓ Resultados salvos em: resultado_teste_final.csv")
    
    # Estatísticas adicionais
    erros_indices = np.where(predictions != y_test)[0]
    print(f"\n9. Estatísticas:")
    print(f"   ✓ Total de amostras: {len(y_test)}")
    print(f"   ✓ Acertos: {(predictions == y_test).sum()}")
    print(f"   ✓ Erros: {len(erros_indices)}")
    print(f"   ✓ Taxa de Erro: {len(erros_indices) / len(y_test):.2%}")
    
    # Mostrar alguns erros
    if len(erros_indices) > 0:
        print(f"\n   Primeiros 5 erros:")
        for i, idx in enumerate(erros_indices[:5]):
            print(f"   - Amostra {idx}: Real={chr(65+y_test[idx])}, Predito={chr(65+predictions[idx])}")
    
    print("\n" + "=" * 70)
    print("TESTE CONCLUÍDO COM SUCESSO!")
    print("=" * 70)


if __name__ == "__main__":
    main()
