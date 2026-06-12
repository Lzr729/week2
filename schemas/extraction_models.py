from typing import Literal, Optional, Union
from pydantic import BaseModel, Field


Number = Optional[Union[int, float]]


class SubscriptionFlow(BaseModel):
    record_type: Literal["subscription_flow"]

    company_code: str
    company_name: str
    pdf_page: Number

    increase_date: Optional[str] = None
    subscriber: Optional[str] = None

    subscribed_shares_wan: Number = None
    subscription_amount_wan: Number = None
    subscription_price_yuan_per_share: Number = None

    evidence_text: Optional[str] = None
    review_notes: Optional[str] = None


class EquitySnapshot(BaseModel):
    record_type: Literal["equity_snapshot"]

    company_code: str
    company_name: str
    pdf_page: Number

    snapshot_time: Optional[str] = None
    equity_scope: Optional[str] = None

    total_shares_wan: Number = None
    total_capital_wan: Number = None

    shareholder_name: Optional[str] = None
    shares_wan: Number = None
    capital_contribution_wan: Number = None
    shareholding_ratio: Optional[str] = None

    evidence_text: Optional[str] = None
    review_notes: Optional[str] = None