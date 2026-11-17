# Histórico da Loja 360

Rotina semanal que monta a linha do tempo de cada loja (treinamentos, racks, exposições especiais, testeiras e roteiro) e publica num site com login.

## Rotina (`tools/atualizar_tudo.py`)

1. Visitas e lojas a partir da exportação de treinamentos.
2. Ativações a partir de três planilhas (rack, exposição especial e testeira), casando a loja pelo nome normalizado.
3. Cobertura de roteiro por colaborador, incluindo lojas do roteiro que ainda não têm visita.
4. **Validação antes de publicar**: confere que o JSON continua válido e que o número de registros não caiu mais de 50%. Se algo falhar, a rotina para sem publicar.

Também há a versão que consulta o banco direto (`tools/montar_historico_loja.ps1`).

## Stack

Python, pandas, PowerShell, SQL Server, Firebase Hosting.

---
Projeto feito no trabalho, publicado sem dados. A página do site ficou de fora porque leva os dados embutidos.
