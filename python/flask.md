
## 如何解决跨域

* 安装第三方库  flask-cors 
* pip install flask-cors 
* 在启动前配置cors

```python
from flask import Flask  
from flask_cors import CORS  
  
app = Flask(__name__)  
CORS(app)  # 允许所有来源的跨域请求  
  
  
@app.route('/')  
def index():  
    return 'flask'  
  
  
if __name__ == '__main__':  
    app.run()
```