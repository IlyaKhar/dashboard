import { useState } from "react";
import { useNavigate, useLocation } from "react-router-dom";
import { Button } from "@/shared/ui/button";
import { ArrowLeft } from "lucide-react";
import type { DeptDrillItem, GroupDrillItem } from "@/shared/api";
import {
  DepartmentsLevel,
  GroupsLevel,
  AnalyticsLevel,
} from "./_components";

type DrillLevel = "departments" | "groups" | "analytics";

export function GroupsPage() {
  const navigate = useNavigate();
  const location = useLocation();

  const routerState = location.state as {
    department?: string;
    departmentData?: DeptDrillItem;
  } | null;

  const [level, setLevel] = useState<DrillLevel>(
    routerState?.departmentData ? "groups" : "departments"
  );
  const [selectedDept, setSelectedDept] = useState<DeptDrillItem | null>(
    routerState?.departmentData ?? null
  );
  const [selectedGroup, setSelectedGroup] = useState<GroupDrillItem | null>(
    null
  );

  function handleSelectDepartment(dept: DeptDrillItem) {
    setSelectedDept(dept);
    setSelectedGroup(null);
    setLevel("groups");
  }

  function handleSelectGroup(group: GroupDrillItem) {
    setSelectedGroup(group);
    setLevel("analytics");
  }

  function handleBack() {
    if (level === "analytics") {
      setSelectedGroup(null);
      setLevel("groups");
    } else if (level === "groups") {
      setSelectedDept(null);
      setLevel("departments");
    } else {
      navigate("/");
    }
  }

  return (
    <div className="space-y-6">
      <Button onClick={handleBack} variant="default" className="h-10 px-4 py-2">
        <ArrowLeft className="mr-2 h-4 w-4" />
        Назад
      </Button>

      {level === "departments" && (
        <DepartmentsLevel onSelectDepartment={handleSelectDepartment} />
      )}
      {level === "groups" && selectedDept && (
        <GroupsLevel department={selectedDept} onSelectGroup={handleSelectGroup} />
      )}
      {level === "analytics" && selectedDept && selectedGroup && (
        <AnalyticsLevel
          department={selectedDept}
          group={selectedGroup}
        />
      )}
    </div>
  );
}
