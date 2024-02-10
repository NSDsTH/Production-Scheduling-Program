# ตัวอย่างข้อมูล Objective Values ของแต่ละโซลูชัน
solutions = [[1, 3], [2, 2], [3, 1], [0, 4], [4, 0]]

# จัดเรียงโซลูชันตาม Objective แต่ละตัวและคำนวณ Crowding Distance

# 1. จัดเรียงโซลูชันตาม Objective
solutions_sorted_by_obj1 = sorted(solutions, key=lambda x: x[0])
solutions_sorted_by_obj2 = sorted(solutions, key=lambda x: x[1])

print(solutions_sorted_by_obj1)
print(solutions_sorted_by_obj2)
# 2. กำหนดค่า Crowding Distance สำหรับโซลูชันที่ขอบเขตเป็นอนันต์และคำนวณระยะห่างสำหรับโซลูชันอื่นๆ
# สร้าง dictionary เพื่อเก็บค่า crowding distance ของแต่ละโซลูชัน
crowding_distances = {tuple(solution): 0 for solution in solutions}

# กำหนดค่าระยะห่างสำหรับโซลูชันที่ขอบเขต (ขอบเขตที่ได้ค่าระยะห่างเป็นอนันต์หรือค่าสูงสุด)
n = len(solutions)
max_distance = float("inf")
crowding_distances[tuple(solutions_sorted_by_obj1[0])] = max_distance
crowding_distances[tuple(solutions_sorted_by_obj1[-1])] = max_distance
crowding_distances[tuple(solutions_sorted_by_obj2[0])] = max_distance
crowding_distances[tuple(solutions_sorted_by_obj2[-1])] = max_distance

# คำนวณระยะห่างสำหรับโซลูชันกลาง
for i in range(1, n - 1):
    distance_obj1 = (
        solutions_sorted_by_obj1[i + 1][0] - solutions_sorted_by_obj1[i - 1][0]
    )
    print(
        "Solution1:",
        solutions_sorted_by_obj1[i],
        solutions_sorted_by_obj1[i + 1],
        solutions_sorted_by_obj1[i - 1],
    )
    distance_obj2 = (
        solutions_sorted_by_obj2[i + 1][1] - solutions_sorted_by_obj2[i - 1][1]
    )
    print(
        "Solution2:",
        solutions_sorted_by_obj2[i],
        solutions_sorted_by_obj2[i + 1],
        solutions_sorted_by_obj2[i - 1],
    )
    # รวมค่าระยะห่างจากทั้งสอง objectives
    crowding_distances[tuple(solutions_sorted_by_obj1[i])] += distance_obj1
    crowding_distances[tuple(solutions_sorted_by_obj2[i])] += distance_obj2

# คืนค่า crowding distances เป็นรายการพร้อมค่า
crowding_distances_sorted = sorted(crowding_distances.items(), key=lambda x: -x[1])

print(crowding_distances_sorted)
