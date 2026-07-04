
from datetime import datetime
from typing import Optional

from pydantic import BaseModel, EmailStr, Field


class MessageResponse(BaseModel):
    success: bool
    message: str


class SendRegisterCodeRequest(BaseModel):
    username: str = Field(min_length=3, max_length=100)
    password: str = Field(min_length=6, max_length=100)
    confirm_password: str = Field(min_length=6, max_length=100)
    phone: str = ""
    email: EmailStr


class VerifyRegisterCodeRequest(BaseModel):
    username: str
    password: str
    confirm_password: str
    phone: str = ""
    email: EmailStr
    code: str
    device_id: Optional[str] = ""
    device_name: Optional[str] = ""


class LoginByAccountRequest(BaseModel):
    account: str
    password: str
    device_id: Optional[str] = ""
    device_name: Optional[str] = ""


class ForgotPasswordRequest(BaseModel):
    account: str


class ResetPasswordRequest(BaseModel):
    account: str
    code: str
    new_password: str = Field(min_length=6, max_length=100)
    confirm_new_password: str = Field(min_length=6, max_length=100)


class AdminGrantFreeRequest(BaseModel):
    account: str
    free_days: int = Field(ge=1, le=3650)
    reason: str = ""


class AdminSetPlanRequest(BaseModel):
    account: str
    plan_type: str  # monthly / quarterly / yearly


class AdminRebindDeviceRequest(BaseModel):
    account: str
    device_id: str
    device_name: str = ""


class AdminDeactivateUserRequest(BaseModel):
    account: str
    is_active: bool


class PaymentReportCreateRequest(BaseModel):
    plan_type: str  # monthly / halfyear / yearly
    amount: Optional[float] = None
    transfer_last5: str = ""
    transfer_time: Optional[datetime] = None
    payer_name: str = ""
    note: str = ""


class FeedbackSubmitRequest(BaseModel):
    topic: str = Field(min_length=1, max_length=100)
    content: str = Field(min_length=1, max_length=5000)
    feedback_id: Optional[str] = ""
    app_version: Optional[str] = "web"
    device_info: Optional[str] = ""
