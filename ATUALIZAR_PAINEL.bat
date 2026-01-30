@echo off
echo =====================================
echo Atualizando painel a partir do Excel
echo =====================================

python python\atualizar_painel_completo.py
python python\atualizar_preco_medio.py

echo Publicando no GitHub...

git add .
git commit -m "Atualização automática KPIs + Preço Médio"
git push

echo =====================================
echo Painel atualizado com sucesso!
echo =====================================

pause
