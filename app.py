from typing import Optional

import httpx
import uvicorn
import msal
from starlette.middleware.sessions import SessionMiddleware
from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi_msal import MSALAuthorization, MSALClientConfig
from fastapi_msal.models import AuthToken


class AppConfig(MSALClientConfig):
    # You can find more Microsoft Graph API endpoints from Graph Explorer
    # https://developer.microsoft.com/en-us/graph/graph-explorer
    endpoint: str = "https://graph.microsoft.com/v1.0/users"  # This resource requires no admin consent
    login_path = "/login"  # default is '/_login_route'
    logout_path = "/logout"  # default is '/_logout_route'


config = AppConfig(_env_file="app_config.env")

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=config.client_credential)
auth = MSALAuthorization(client_config=config)
app.include_router(auth.router)

templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    token: Optional[AuthToken] = await auth.get_session_token(request=request)
    if not token or not token.id_token_claims:
        return RedirectResponse(url=config.login_path)
    context = {
        "request": request,
        "user_name": token.id_token_claims,
        "version": msal.__version__,
    }
    return templates.TemplateResponse(name="index.html", context=context)


@app.get("/graphcall")
async def graphcall(request: Request):
    token: Optional[AuthToken] = await auth.handler.get_token_from_session(request=request)
    if not token or not token.access_token:
        return RedirectResponse(url=config.login_path)
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            config.endpoint, headers={"Authorization": "Bearer " + token.access_token},
        )
    graph_data = resp.json()
    context = {"request": request, "result": graph_data}
    return templates.TemplateResponse(name="display.html", context=context)


if __name__ == "__main__":
    uvicorn.run("app:app", host="localhost", port=5000, reload=True)
