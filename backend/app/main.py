import logging
from fastapi import FastAPI
from app.api.endpoints import router as api_router
from app.api.hyperlink_routes import router as hyperlink_router
from fastapi.middleware.cors import CORSMiddleware

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = FastAPI(title="Inspector: SVN CSV + Review Checklist (with advanced compare)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(api_router, prefix="/api")
app.include_router(hyperlink_router, prefix="/api/hyperlinks", tags=["Hyperlinks"])

@app.get("/health")
async def health():
    logger.info("Health check endpoint called.")
    return {"status": "ok"}
