from fastapi import HTTPException, Security, status
from fastapi.security.api_key import APIKeyHeader

from .. import settings


async def authenticate(api_key: str = Security(APIKeyHeader(name="Authorization"))):
    """Validate the API_KEY as delivered by the Authorization header"""
    if api_key != settings.XLWINGS_API_KEY:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid API Key",
        )
