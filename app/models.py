from pydantic import BaseModel, Field


class GenerateRequest(BaseModel):
    topic: str = Field(..., min_length=1, max_length=500)
    num_slides: int = Field(default=8, ge=3, le=20)
    style: str = Field(default="business")
    additional_instructions: str = Field(default="", max_length=2000)
    language: str = Field(default="ja")
    pdf_text: str = Field(default="")
    csv_analysis: dict | None = Field(default=None)


class ChartData(BaseModel):
    chart_type: str = "bar"
    categories: list[str] = []
    series: list[dict] = []


class StatItem(BaseModel):
    value: str
    label: str


class SlideData(BaseModel):
    slide_number: int
    layout: str
    title: str
    subtitle: str = ""
    bullet_points: list[str] = []
    image_keyword: str = ""
    chart: ChartData | None = None
    stats: list[StatItem] = []
    speaker_notes: str = ""


class PresentationData(BaseModel):
    title: str
    theme: str = "ocean_blue"
    slides: list[SlideData]
