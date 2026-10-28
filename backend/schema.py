"""Table definitions shared by the Excel importer and the API.
Column names are identical to the workbook headers, so rows can be sent to the frontend unchanged."""

DEPTS = ["AI/ML", "Data Analytics", "Finance", "HR", "IT Support", "Marketing", "Operations", "Product", "QA", "Research", "Sales", "Software Development"]
PRIORITY = ["Critical", "High", "Medium", "Low"]
CTYPES = ["Availability", "Deadline", "Department Match", "Maximum Workload", "Minimum Experience", "Required Skill",
          "Resource Availability", "Task Dependency", "Team Capacity", "Working Hours"]


def T(table, pk, cols, route_fk=None, enums=None, ranges=None, required=None, kinds=None):
    return dict(table=table, pk=pk, cols=cols, fks=route_fk or {}, enums=enums or {}, ranges=ranges or {},
                required=required, kinds=kinds or {})


# cols: name -> text | int | real ; kinds: date | datetime ; fks: column -> (route, pk column)
TABLES = {
    "employees": T("Employees", "Employee_ID", dict(
        Employee_ID="text", Employee_Name="text", Department="text", Role="text", Skills="text", Experience_Years="int",
        Availability_Hours_Per_Day="int", Current_Workload_Percent="int", Performance_Score="int", Team_ID="text", Work_Mode="text"),
        enums=dict(Work_Mode=["Hybrid", "On-site", "Remote"]),
        ranges=dict(Experience_Years=(0, 60), Availability_Hours_Per_Day=(0, 24), Current_Workload_Percent=(0, 100), Performance_Score=(0, 100))),
    "departments": T("Departments", "Department_ID", dict(
        Department_ID="text", Department_Name="text", Typical_Team_Size="int", Department_Type="text"),
        enums=dict(Department_Type=["Business", "Support", "Technical"]), ranges=dict(Typical_Team_Size=(1, 1000))),
    "projects": T("Projects", "Project_ID", dict(
        Project_ID="text", Project_Name="text", Primary_Department="text", Priority="text", Status="text", Team_Size="int", Target_Completion_Date="text"),
        enums=dict(Priority=PRIORITY, Status=["Active", "Completed", "On Hold", "Planning"]), ranges=dict(Team_Size=(1, 1000)),
        kinds=dict(Target_Completion_Date="date")),
    "tasks": T("Tasks", "Task_ID", dict(
        Task_ID="text", Project_ID="text", Task_Name="text", Required_Skills="text", Priority="text", Estimated_Hours="int",
        Deadline="text", Dependency_Task_ID="text", Required_Employees="int", Status="text", Department="text"),
        route_fk=dict(Project_ID=("projects", "Project_ID")),
        enums=dict(Priority=PRIORITY, Status=["Pending", "In Progress", "Completed", "Blocked"]),
        ranges=dict(Estimated_Hours=(1, 10000), Required_Employees=(1, 100)), kinds=dict(Deadline="date"),
        required=["Project_ID", "Task_Name", "Required_Skills", "Priority", "Estimated_Hours", "Deadline", "Required_Employees", "Status", "Department"]),
    "assignments": T("Assignments", "Assignment_ID", dict(
        Assignment_ID="text", Employee_ID="text", Task_ID="text", Skill_Match_Score="real", Availability_Score="real",
        Workload_Factor="real", Assignment_Score="real", Assignment_Status="text"),
        route_fk=dict(Employee_ID=("employees", "Employee_ID"), Task_ID=("tasks", "Task_ID")),
        enums=dict(Assignment_Status=["Assigned", "Completed", "Rejected", "Suggested"]),
        ranges=dict(Skill_Match_Score=(0, 1), Availability_Score=(0, 1), Workload_Factor=(0, 1), Assignment_Score=(0, 100))),
    "constraints": T("Constraints", "Constraint_ID", dict(
        Constraint_ID="text", Task_ID="text", Constraint_Type="text", Constraint_Value="text", Constraint_Level="text"),
        route_fk=dict(Task_ID=("tasks", "Task_ID")),
        enums=dict(Constraint_Type=CTYPES, Constraint_Level=["Hard", "Soft"])),
    "communications": T("Communications", "Message_ID", dict(
        Message_ID="text", Sender_ID="text", Receiver_ID="text", Project_ID="text", Communication_Type="text",
        Message_Text="text", Timestamp="text", Priority="text", Action_Required="text"),
        route_fk=dict(Sender_ID=("employees", "Employee_ID"), Receiver_ID=("employees", "Employee_ID"), Project_ID=("projects", "Project_ID")),
        enums=dict(Communication_Type=["Announcement", "Chat", "Document Comment", "Email", "Project Update"],
                   Priority=["High", "Medium", "Low"], Action_Required=["Yes", "No"]), kinds=dict(Timestamp="datetime")),
    "meetings": T("Meetings", "Meeting_ID", dict(
        Meeting_ID="text", Project_ID="text", Participants="text", Topic="text", Discussion_Summary="text",
        Decisions="text", Action_Items="text", Meeting_Date="text"),
        route_fk=dict(Project_ID=("projects", "Project_ID")), kinds=dict(Meeting_Date="date")),
    "resources": T("Resources", "Resource_ID", dict(
        Resource_ID="text", Resource_Name="text", Resource_Type="text", Capacity="int", Utilization_Percent="int",
        Availability_Status="text", Project_ID="text"),
        route_fk=dict(Project_ID=("projects", "Project_ID")),
        enums=dict(Resource_Type=["Facility", "Hardware", "Service", "Software"],
                   Availability_Status=["Available", "Partially Available", "Unavailable"]),
        ranges=dict(Capacity=(0, 100000), Utilization_Percent=(0, 100))),
    "dependencies": T("Task_Dependencies", "Dependency_ID", dict(
        Dependency_ID="text", Predecessor_Task_ID="text", Successor_Task_ID="text", Dependency_Type="text"),
        route_fk=dict(Predecessor_Task_ID=("tasks", "Task_ID"), Successor_Task_ID=("tasks", "Task_ID")),
        enums=dict(Dependency_Type=["Finish-to-Start", "Start-to-Start", "Finish-to-Finish", "Start-to-Finish"])),
}
for _r, _s in TABLES.items():  # primary key is generated when omitted; everything else is required unless listed otherwise
    if _s["required"] is None:
        _s["required"] = [c for c in _s["cols"] if c != _s["pk"]]
