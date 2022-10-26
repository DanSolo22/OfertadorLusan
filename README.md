# Ofertador Lusan

Programa para convertir los csv de lusan en .docx con buen formato y diseño llamativo.

# Configuración

Para crear un entorno virtual:

### Desde el terminal

```bash
cd C:\Users\DanSolo\PycharmProjects\OfertadorLusan
python -m venv venv
cd venv
Scripts\activate
pip install --upgrade pip
pip install django
pip install pywin32
pip install jinja2
pip install mysqlclient
pip install PyMySql
pip install python-dotenv
pip install wheel
pip install docxtpl
pip install python-docx
pip install docx
cd ..
manage.py migrate
manage.py collectstatic
```