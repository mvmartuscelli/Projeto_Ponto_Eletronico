import tkcap
from sistema_ponto_visual import AppPonto

def capture_and_close():
    cap = tkcap.CAP(app)
    cap.capture("verification.png")
    app.destroy()

app = AppPonto()
app.after(2000, capture_and_close)  # Wait 2 seconds for the window to render
app.mainloop()
