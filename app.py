from app import create_app, db

app = create_app()
app.run(debug=True)

# 将db、user、post注册进shell,通过flask shell访问
@app.shell_context_processor
def make_shell_context():
    return {'db': db}
