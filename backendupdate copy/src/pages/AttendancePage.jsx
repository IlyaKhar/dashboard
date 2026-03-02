import { useState, useEffect, useMemo } from "react";
import { useNavigate } from "react-router-dom";
import {
  BarChart,
  Bar,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from "recharts";
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
import { Separator } from "@/components/ui/separator";
import { Users, Calendar, AlertTriangle } from "lucide-react";
import loadAttendance from "@/utils/xlsxAttendanceConverter";

const COLORS = [
  "#000000",
  "#1a1a1a",
  "#333333",
  "#4d4d4d",
  "#666666",
  "#808080",
];

export function AttendancePage() {
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

  const stats = useMemo(() => {
    if (!data.length) return null;

    const allStudents = [];
    const missedByDate = {};
    const missedByGroup = {};
    const missedByDepartment = {};

    data.forEach((dept) => {
      missedByDepartment[dept.department] = 0;
      dept.groups?.forEach((group) => {
        missedByGroup[group.group] = 0;
        group.students?.forEach((student) => {
          const totalMissed =
            student.attendance?.reduce((sum, att) => sum + att.missed, 0) || 0;
          allStudents.push({
            name: student.student,
            group: group.group,
            department: dept.department,
            totalMissed,
            attendance: student.attendance || [],
          });
          missedByGroup[group.group] += totalMissed;
          missedByDepartment[dept.department] += totalMissed;

          student.attendance?.forEach((att) => {
            missedByDate[att.date] = (missedByDate[att.date] || 0) + att.missed;
          });
        });
      });
    });

    const totalMissed = allStudents.reduce((sum, s) => sum + s.totalMissed, 0);

    const topOffenders = [...allStudents]
      .sort((a, b) => b.totalMissed - a.totalMissed)
      .slice(0, 10);

    const dateData = Object.entries(missedByDate)
      .sort(([a], [b]) => new Date(a) - new Date(b))
      .map(([date, missed]) => ({
        date: new Date(date).toLocaleDateString("ru-RU", {
          day: "2-digit",
          month: "2-digit",
        }),
        missed,
      }));

    const groupData = Object.entries(missedByGroup)
      .map(([group, missed]) => ({ group, missed }))
      .sort((a, b) => b.missed - a.missed)
      .slice(0, 10);

    const deptData = Object.entries(missedByDepartment).map(
      ([department, missed]) => ({
        department:
          department.length > 30
            ? department.substring(0, 30) + "..."
            : department,
        missed,
      })
    );

    return {
      totalStudents: allStudents.length,
      totalMissed,
      topOffenders,
      dateData,
      groupData,
      deptData,
      allStudents,
    };
  }, [data]);

  if (loading) {
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-lg">Загрузка данных...</div>
      </div>
    );
  }

  if (!stats) {
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-lg">Нет данных для отображения</div>
      </div>
    );
  }

  return (
    <div className="space-y-8">
      {/* Статистические карточки */}
      <div className="space-y-4">
        <div>
          <h2 className="text-lg font-semibold text-black">Общая статистика</h2>
          <p className="text-sm text-muted-foreground">
            Основные показатели посещаемости
          </p>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <Card>
            <CardHeader className="flex flex-row items-center justify-between pb-2 space-y-0">
              <CardTitle className="text-sm font-medium text-black">
                Всего студентов
              </CardTitle>
              <Users className="h-4 w-4 text-black" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-black">
                {stats.totalStudents}
              </div>
              <p className="text-xs text-gray-600 mt-1">с пропусками</p>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between pb-2 space-y-0">
              <CardTitle className="text-sm font-medium text-gray-700">
                Всего пропусков
              </CardTitle>
              <Calendar className="h-4 w-4 text-gray-600" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-black">
                {stats.totalMissed}
              </div>
              <p className="text-xs text-gray-600 mt-1">часов</p>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between pb-2 space-y-0">
              <CardTitle className="text-sm font-medium text-black">
                Критические случаи
              </CardTitle>
              <AlertTriangle className="h-4 w-4 text-black" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-black">
                {stats.allStudents.filter((s) => s.totalMissed > 20).length}
              </div>
              <p className="text-xs text-gray-600 mt-1">
                &gt;20 часов пропусков
              </p>
            </CardContent>
          </Card>
        </div>
      </div>

      <Separator />

      {/* Графики */}
      <div className="space-y-4">
        <div>
          <h2 className="text-lg font-semibold text-black">Аналитика</h2>
          <p className="text-sm text-muted-foreground">
            Динамика и распределение пропусков
          </p>
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <Card>
            <CardHeader>
              <CardTitle>Динамика пропусков по датам</CardTitle>
              <CardDescription>
                Количество пропущенных часов по дням
              </CardDescription>
            </CardHeader>
            <CardContent>
              <ResponsiveContainer width="100%" height={300}>
                <LineChart data={stats.dateData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                  <XAxis dataKey="date" stroke="#64748b" fontSize={12} />
                  <YAxis stroke="#64748b" fontSize={12} />
                  <Tooltip
                    contentStyle={{
                      backgroundColor: "white",
                      border: "1px solid #e5e7eb",
                      borderRadius: "8px",
                    }}
                  />
                  <Line
                    type="monotone"
                    dataKey="missed"
                    stroke="#000000"
                    strokeWidth={2}
                    dot={{ fill: "#000000", r: 4 }}
                    activeDot={{ r: 6 }}
                    name="Пропущено часов"
                  />
                </LineChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Топ-10 групп по пропускам</CardTitle>
              <CardDescription>
                Группы с наибольшим количеством пропусков
              </CardDescription>
            </CardHeader>
            <CardContent>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={stats.groupData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                  <XAxis dataKey="group" stroke="#64748b" fontSize={12} />
                  <YAxis stroke="#64748b" fontSize={12} />
                  <Tooltip
                    contentStyle={{
                      backgroundColor: "white",
                      border: "1px solid #e5e7eb",
                      borderRadius: "8px",
                    }}
                  />
                  <Bar
                    dataKey="missed"
                    fill="#000000"
                    radius={[8, 8, 0, 0]}
                    name="Пропущено часов"
                    onClick={(data) => navigate(`/group/${data.group}`)}
                    cursor="pointer"
                  />
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className="lg:col-span-2">
            <CardHeader>
              <CardTitle>Распределение пропусков по отделениям</CardTitle>
              <CardDescription>
                Общее количество пропусков по отделениям
              </CardDescription>
            </CardHeader>
            <CardContent>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <ResponsiveContainer width="100%" height={300}>
                  <PieChart>
                    <Pie
                      data={stats.deptData}
                      cx="50%"
                      cy="50%"
                      labelLine={false}
                      label={({ department, percent }) =>
                        `${
                          department.split(" ")[1] ||
                          department.substring(0, 15)
                        } (${(percent * 100).toFixed(0)}%)`
                      }
                      outerRadius={100}
                      fill="#8884d8"
                      dataKey="missed"
                    >
                      {stats.deptData.map((entry, index) => (
                        <Cell
                          key={`cell-${index}`}
                          fill={COLORS[index % COLORS.length]}
                        />
                      ))}
                    </Pie>
                    <Tooltip />
                  </PieChart>
                </ResponsiveContainer>
                <div className="flex flex-col justify-center space-y-3">
                  {stats.deptData.map((dept, index) => (
                    <div key={index} className="flex items-center gap-3">
                      <div
                        className="w-4 h-4 rounded"
                        style={{
                          backgroundColor: COLORS[index % COLORS.length],
                        }}
                      />
                      <div className="flex-1">
                        <div className="text-sm font-medium text-black">
                          {dept.department}
                        </div>
                        <div className="text-xs text-gray-600">
                          {dept.missed} часов
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>

      <Separator />

      {/* Таблица топ-10 */}
      <div className="space-y-4">
        <div>
          <h2 className="text-lg font-semibold text-black">Рейтинги</h2>
          <p className="text-sm text-muted-foreground">
            Студенты и группы с наибольшим количеством пропусков
          </p>
        </div>
        <Card>
          <CardHeader>
            <CardTitle>Топ-10 студентов по количеству пропусков</CardTitle>
            <CardDescription>
              Студенты с наибольшим количеством пропущенных часов
            </CardDescription>
          </CardHeader>
          <CardContent>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="w-12">#</TableHead>
                  <TableHead>ФИО студента</TableHead>
                  <TableHead>Группа</TableHead>
                  <TableHead>Отделение</TableHead>
                  <TableHead className="text-right">Пропущено часов</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {stats.topOffenders.map((student, index) => (
                  <TableRow key={index}>
                    <TableCell className="font-medium">
                      <div className="flex items-center justify-center w-8 h-8 rounded-full bg-black text-white text-sm font-semibold">
                        {index + 1}
                      </div>
                    </TableCell>
                    <TableCell className="font-medium">
                      {student.name}
                    </TableCell>
                    <TableCell>
                      <Badge variant="secondary">{student.group}</Badge>
                    </TableCell>
                    <TableCell className="text-sm text-gray-700">
                      {student.department.length > 40
                        ? student.department.substring(0, 40) + "..."
                        : student.department}
                    </TableCell>
                    <TableCell className="text-right">
                      <Badge
                        variant={
                          student.totalMissed > 20
                            ? "default"
                            : student.totalMissed > 10
                            ? "secondary"
                            : "outline"
                        }
                        className={
                          student.totalMissed > 20
                            ? "bg-black"
                            : student.totalMissed > 10
                            ? "bg-gray-700 text-white"
                            : ""
                        }
                      >
                        {student.totalMissed} ч.
                      </Badge>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      </div>

      <Separator />

      {/* Полный список */}
      <div className="space-y-4">
        <div>
          <h2 className="text-lg font-semibold text-black">Полный реестр</h2>
          <p className="text-sm text-muted-foreground">
            Все студенты с зафиксированными пропусками
          </p>
        </div>
        <Card>
          <CardHeader>
            <CardTitle>Полный список студентов с пропусками</CardTitle>
            <CardDescription>
              Все студенты, имеющие пропуски по неуважительной причине
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="max-h-96 overflow-y-auto">
              <Table>
                <TableHeader className="sticky top-0 bg-white">
                  <TableRow>
                    <TableHead>ФИО студента</TableHead>
                    <TableHead>Группа</TableHead>
                    <TableHead>Отделение</TableHead>
                    <TableHead className="text-right">
                      Всего пропусков
                    </TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {stats.allStudents
                    .sort((a, b) => b.totalMissed - a.totalMissed)
                    .map((student, index) => (
                      <TableRow key={index}>
                        <TableCell className="font-medium">
                          {student.name}
                        </TableCell>
                        <TableCell>
                          <Badge variant="outline">{student.group}</Badge>
                        </TableCell>
                        <TableCell className="text-sm text-gray-700">
                          {student.department.length > 40
                            ? student.department.substring(0, 40) + "..."
                            : student.department}
                        </TableCell>
                        <TableCell className="text-right">
                          <Badge
                            variant={
                              student.totalMissed > 20
                                ? "default"
                                : student.totalMissed > 10
                                ? "secondary"
                                : "outline"
                            }
                            className={
                              student.totalMissed > 20
                                ? "bg-black"
                                : student.totalMissed > 10
                                ? "bg-gray-700 text-white"
                                : ""
                            }
                          >
                            {student.totalMissed} ч.
                          </Badge>
                        </TableCell>
                      </TableRow>
                    ))}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
