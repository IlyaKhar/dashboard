import { useState, useEffect, useMemo } from "react";
import { useParams, useNavigate } from "react-router-dom";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { ArrowLeft, Users, Calendar } from "lucide-react";
import loadAttendance from "@/utils/xlsxAttendanceConverter";

export function GroupPage() {
  const { groupName } = useParams();
  const navigate = useNavigate();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    loadAttendance()
      .then((jsonData) => {
        setData(jsonData);
        setLoading(false);
      })
      .catch((err) => {
        console.error("Error loading data:", err);
        setLoading(false);
      });
  }, []);

  const groupData = useMemo(() => {
    if (!data.length || !groupName) return null;

    let groupInfo = null;
    let departmentName = "";

    // Находим группу
    for (const dept of data) {
      const foundGroup = dept.groups?.find((g) => g.group === groupName);
      if (foundGroup) {
        groupInfo = foundGroup;
        departmentName = dept.department;
        break;
      }
    }

    if (!groupInfo) return null;

    // Собираем студентов с пропусками
    const students = groupInfo.students
      ?.map((student) => {
        const totalMissed =
          student.attendance?.reduce((sum, att) => sum + att.missed, 0) || 0;
        return {
          name: student.student,
          totalMissed,
          attendance: student.attendance || [],
        };
      })
      .filter((s) => s.totalMissed > 0)
      .sort((a, b) => b.totalMissed - a.totalMissed);

    const totalMissed =
      students?.reduce((sum, s) => sum + s.totalMissed, 0) || 0;

    return {
      groupName,
      departmentName,
      students,
      totalMissed,
      totalStudents: students?.length || 0,
    };
  }, [data, groupName]);

  const getBadgeVariant = (hours) => {
    if (hours > 20) return "default";
    if (hours > 10) return "secondary";
    return "outline";
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-white flex items-center justify-center">
        <p className="text-lg text-black">Загрузка данных...</p>
      </div>
    );
  }

  if (!groupData) {
    return (
      <div className="min-h-screen bg-white flex items-center justify-center">
        <div className="text-center space-y-4">
          <p className="text-lg text-black">Группа не найдена</p>
          <Button onClick={() => navigate("/")}>
            <ArrowLeft className="h-4 w-4 mr-2" />
            Вернуться назад
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Шапка */}
      <div className="space-y-4">
        <Button
          variant="outline"
          onClick={() => navigate("/")}
          className="mb-4"
        >
          <ArrowLeft className="h-4 w-4 mr-2" />
          Назад к общей статистике
        </Button>

        <div>
          <h1 className="text-3xl font-semibold tracking-tight text-black">
            Группа {groupData.groupName}
          </h1>
          <p className="text-sm text-muted-foreground mt-1">
            {groupData.departmentName}
          </p>
        </div>
      </div>

      {/* Статистика группы */}
      <div className="grid gap-6 md:grid-cols-2">
        <Card>
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
            <CardTitle className="text-sm font-medium text-black">
              Студентов с пропусками
            </CardTitle>
            <Users className="h-4 w-4 text-black" />
          </CardHeader>
          <CardContent>
            <div className="text-2xl font-bold text-black">
              {groupData.totalStudents}
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
            <CardTitle className="text-sm font-medium text-black">
              Всего пропущено
            </CardTitle>
            <Calendar className="h-4 w-4 text-black" />
          </CardHeader>
          <CardContent>
            <div className="text-2xl font-bold text-black">
              {groupData.totalMissed} ч
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Таблица студентов */}
      <Card>
        <CardHeader>
          <CardTitle>Список студентов с пропусками</CardTitle>
          <CardDescription>
            Обучающиеся группы {groupData.groupName} с пропусками по
            неуважительной причине
          </CardDescription>
        </CardHeader>
        <CardContent>
          {groupData.students && groupData.students.length > 0 ? (
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="text-black">#</TableHead>
                  <TableHead className="text-black">ФИО студента</TableHead>
                  <TableHead className="text-right text-black">
                    Пропущено часов
                  </TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {groupData.students.map((student, idx) => (
                  <TableRow key={idx}>
                    <TableCell className="font-medium text-black">
                      {idx + 1}
                    </TableCell>
                    <TableCell className="text-black">{student.name}</TableCell>
                    <TableCell className="text-right">
                      <Badge variant={getBadgeVariant(student.totalMissed)}>
                        {student.totalMissed} ч
                      </Badge>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          ) : (
            <p className="text-center text-muted-foreground py-8">
              В данной группе нет студентов с пропусками
            </p>
          )}
        </CardContent>
      </Card>
    </div>
  );
}
