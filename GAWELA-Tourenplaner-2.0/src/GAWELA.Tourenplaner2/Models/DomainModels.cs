namespace GAWELA.Tourenplaner2.Models;

public sealed class Employee
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = string.Empty;
    public string Short { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public bool Active { get; set; } = true;
    public string CreatedAt { get; set; } = DateTime.Now.ToString("s");
}

public sealed class LoadingArea
{
    public int LengthCm { get; set; }
    public int WidthCm { get; set; }
    public int HeightCm { get; set; }
}

public sealed class Vehicle
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Type { get; set; } = "other";
    public string Name { get; set; } = string.Empty;
    public string LicensePlate { get; set; } = string.Empty;
    public int MaxPayloadKg { get; set; }
    public int MaxTrailerLoadKg { get; set; }
    public bool Active { get; set; } = true;
    public string Notes { get; set; } = string.Empty;
    public int VolumeM3 { get; set; }
    public LoadingArea? LoadingArea { get; set; }
    public string CreatedAt { get; set; } = DateTime.Now.ToString("s");
    public string UpdatedAt { get; set; } = DateTime.Now.ToString("s");
}

public sealed class Trailer
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = string.Empty;
    public string LicensePlate { get; set; } = string.Empty;
    public int MaxPayloadKg { get; set; }
    public bool Active { get; set; } = true;
    public string Notes { get; set; } = string.Empty;
    public int VolumeM3 { get; set; }
    public LoadingArea? LoadingArea { get; set; }
    public string CreatedAt { get; set; } = DateTime.Now.ToString("s");
    public string UpdatedAt { get; set; } = DateTime.Now.ToString("s");
}

public sealed class VehiclePayload
{
    public List<Vehicle> Vehicles { get; set; } = [];
    public List<Trailer> Trailers { get; set; } = [];
}

public sealed class TourStop
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public double? Lat { get; set; }
    public double? Lon { get; set; }
    public int Order { get; set; }
    public string TimeWindowStart { get; set; } = string.Empty;
    public string TimeWindowEnd { get; set; } = string.Empty;
    public int ServiceMinutes { get; set; }
    public string PlannedArrival { get; set; } = string.Empty;
    public string PlannedDeparture { get; set; } = string.Empty;
    public bool ScheduleConflict { get; set; }
    public string ScheduleConflictText { get; set; } = string.Empty;
    public int WaitMinutes { get; set; }
}

public sealed class Tour
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Date { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public List<TourStop> Stops { get; set; } = [];
    public List<string> EmployeeIds { get; set; } = [];
    public string StartTime { get; set; } = "08:00";
    public string RouteMode { get; set; } = "car";
    public string? VehicleId { get; set; }
    public string? TrailerId { get; set; }
    public Dictionary<string, int> TravelTimeCache { get; set; } = [];
}

public sealed class SegmentDetail
{
    public int Index { get; set; }
    public int? TravelMinutes { get; set; }
    public string Arrival { get; set; } = string.Empty;
    public string Departure { get; set; } = string.Empty;
    public int WaitMinutes { get; set; }
}

public sealed class ScheduleResult
{
    public List<TourStop> Stops { get; set; } = [];
    public int TotalTravelMinutes { get; set; }
    public int TotalServiceMinutes { get; set; }
    public int TotalWaitMinutes { get; set; }
    public string EndTime { get; set; } = string.Empty;
    public List<SegmentDetail> SegmentDetails { get; set; } = [];
    public bool HasConflicts { get; set; }
}
