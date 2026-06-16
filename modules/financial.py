def calculate_financial_impact(
    total_wait_minutes,
    improvement_rate,
    hourly_labor_cost,
    working_days_per_month
):
    daily_saved_minutes = total_wait_minutes * (improvement_rate / 100)
    monthly_saved_hours = (daily_saved_minutes * working_days_per_month) / 60
    yearly_saved_hours = monthly_saved_hours * 12
    yearly_saving = yearly_saved_hours * hourly_labor_cost

    estimated_operational_loss = yearly_saving / 0.35 if yearly_saving > 0 else 0
    estimated_project_cost = yearly_saving * 0.20 if yearly_saving > 0 else 0

    roi = (yearly_saving / estimated_project_cost) * 100 if estimated_project_cost > 0 else 0
    payback_month = estimated_project_cost / (yearly_saving / 12) if yearly_saving > 0 else 0

    return {
        "Günlük Kazanılan Süre (dk)": round(daily_saved_minutes, 2),
        "Aylık Kazanılan Süre (saat)": round(monthly_saved_hours, 2),
        "Yıllık Kazanılan Süre (saat)": round(yearly_saved_hours, 2),
        "Tahmini Yıllık Tasarruf": round(yearly_saving, 2),
        "Toplam Operasyonel Kayıp": round(estimated_operational_loss, 2),
        "İyileştirme Potansiyeli": round(yearly_saving, 2),
        "ROI (%)": round(roi, 2),
        "Geri Dönüş Süresi (Ay)": round(payback_month, 2)
    }