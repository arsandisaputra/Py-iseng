import os
import sys
import openpyxl
import win32com.client
import subprocess

import random

batasBawah = 75
batasAtas = 95


RPH = 0
PTS = 0
PAS = 0

def open_excel_file(file_path):
    try:
        subprocess.Popen(['start', 'excel', file_path], shell=True)
        print(f'File {file_path} dibuka dengan Excel.')
    except Exception as e:
        print(f'Error: {e}')


def close_excel_file(file_path):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        workbook.Close(SaveChanges=False)
        excel.Quit()
        print(f'File {file_path} ditutup.')
    except Exception as e:
        print(f'Error: {e}')


def acakByMean(mean):
    var1 = mean - batasBawah
    var2 = batasAtas - mean
    randomVar = min(var1, var2)
    # print('randomVar', randomVar)
    angka = random.randint(-1 * randomVar, randomVar)
    return angka


def GenerateNilai(mean, mean2):
    # Ubah nilai sel
    while True:
        acak = acakByMean(mean)
        # print(acak)
        RPH = mean + acak - (acak // 2)
        if RPH >= batasBawah and RPH <= batasAtas:
            break

    # print('RPH', RPH)
    

    sisa = (mean * 4) - ((RPH * 2) + (mean*2))
    # print(sisa)

    while True:
        randPosisi = random.randint(0, 1)
        # print('pos', randPosisi)
        randNext = acakByMean(mean) // 2
        sisaLast = sisa - (sisa // 2)
        if randPosisi == 0:
            PTS = mean + randNext + (sisa // 2)
            PAS = mean - randNext + (sisaLast)
        else:
            PTS = mean - randNext + (sisa // 2)
            PAS = mean + randNext + (sisaLast)

        if PTS >= batasBawah and PTS <= batasAtas and PAS >= batasBawah and PAS <= batasAtas:
            break

    # print('Mean awal:', mean, ' Mean kemudian:', ((RPH*2) + PTS + PAS)/4)

    P = [RPH, RPH, RPH, RPH]
    sisaP = sum(P)
    pos = [0, 1, 2, 3]
    randPosisi = random.randint(0, len(pos)-1)
    # print(randPosisi)
    # print(pos[randPosisi])
    index = pos[randPosisi]

    while True:
        acak = acakByMean(RPH)
        # print(acak)
        val = RPH + acak
        if val >= batasBawah and val <= batasAtas:
            P[index] = val
            sisaP -= val
            break

    pos.pop(randPosisi)
    # print(pos)

    randPosisi = random.randint(0, len(pos)-1)
    # print(randPosisi)
    # print(pos[randPosisi])
    index = pos[randPosisi]
    while True:
        acak = acakByMean(RPH)
        # print(acak)
        val = RPH + acak
        if acak != 0 and val > batasBawah and val < batasAtas:
            P[index] = val
            sisaP -= val
            break

    pos.pop(randPosisi)
    # print(pos)

    index1 = pos[0]
    index2 = pos[1]
    sisaP -= (RPH * 2)
    # print("sisa", sisaP)
    while True:
        randPosisi = random.randint(0, 1)
        # print('pos', randPosisi)
        randNext = acakByMean(RPH) // 2
        sisaPLast = sisaP - (sisaP // 2)
        if randPosisi == 0:
            P[index1] = RPH + randNext + (sisaP // 2)
            P[index2] = RPH - randNext + (sisaPLast)
        else:
            P[index1] = RPH - randNext + (sisaP // 2)
            P[index2] = RPH + randNext + (sisaPLast)

        if P[index1] >= batasBawah and P[index1] <= batasAtas and P[index2] >= batasBawah and P[index2] <= batasAtas:
            break
    P[index1] = P[index1]
    P[index2] = P[index2]
    # print(P)

    # print('mean RPH', sum(P) / len(P))
    # P1 = 0
    # P2 = 0
    # P3 = 0
    # P4 = 0

    # NilaiPengetahuan = mean
    print('-------')

    K = [mean2, mean2, mean2, mean2]
    sisaK = sum(K)
    pos = [0, 1, 2, 3]
    randPosisi = random.randint(0, len(pos)-1)
    # print(randPosisi)
    # print(pos[randPosisi])
    index = pos[randPosisi]

    while True:
        acak = acakByMean(mean2)
        # print(acak)
        val = mean2 + acak
        if val >= batasBawah and val <= batasAtas:
            K[index] = val
            sisaK -= val
            break

    pos.pop(randPosisi)
    # print(pos)

    randPosisi = random.randint(0, len(pos)-1)
    # print(randPosisi)
    # print(pos[randPosisi])
    index = pos[randPosisi]
    while True:
        acak = acakByMean(mean2)
        # print(acak)
        val = mean2 + acak
        if acak != 0 and val > batasBawah and val < batasAtas:
            K[index] = val
            sisaK -= val
            break

    pos.pop(randPosisi)
    # print(pos)

    index1 = pos[0]
    index2 = pos[1]
    sisaK -= (mean2 * 2)
    # print("sisa", sisaK)
    while True:
        randPosisi = random.randint(0, 1)
        # print('pos', randPosisi)
        randNext = acakByMean(mean2) // 2
        sisaKLast = sisaK - (sisaK // 2)
        if randPosisi == 0:
            K[index1] = mean2 + randNext + (sisaK // 2)
            K[index2] = mean2 - randNext + (sisaKLast)
        else:
            K[index1] = mean2 - randNext + (sisaK // 2)
            K[index2] = mean2 + randNext + (sisaKLast)

        if K[index1] >= batasBawah and K[index1] <= batasAtas and K[index2] >= batasBawah and K[index2] <= batasAtas:
            break
    K[index1] = K[index1]
    K[index2] = K[index2]

    file_path = "nilai.txt"
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write("Nilai P1 - P4:\n")
        nilaiP = "\t".join(list(map(str, P)))
        nilaiP += "\n"
        file.write(nilaiP)
        file.write("PTS\tPAS:\n")
        nilai = str(PTS) + "\t" + str(PAS)
        nilai += "\n"
        file.write(nilai)
        file.write("Nilai K1 - K4:\n")
        nilaiK = "\t".join(list(map(str, K)))
        nilaiK += "\n"
        file.write(nilaiK)
    # print(K)
    # print('mean2', sum(K) / len(K))


def save_to_excel():
    # Buka file Excel
    excel_file = 'data_excel.xlsx'
    open_excel_file(excel_file)
    close_excel_file(excel_file)

    workbook = openpyxl.Workbook()

    # Pilih lembar kerja (sheet) yang ingin diubah
    sheet = workbook.active

    sheet['A5'] = P1
    sheet['B5'] = P2
    sheet['C5'] = P2
    sheet['D5'] = P2

    sheet['F5'] = RPH
    sheet['G5'] = PTS
    sheet['H5'] = PAS

    # Simpan perubahan
    workbook.save(excel_file)

    print('Perubahan file', excel_file, 'telah disimpan.')


# script.py


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python GenerateNilaiByAverage.py <integer_parameter> <integer_parameter>")
        sys.exit(1)

    try:
        mean = int(sys.argv[1])
        mean2 = int(sys.argv[2])
        GenerateNilai(mean, mean2)
        # print(f"The provided integer parameter is: {integer_parameter}")

    except ValueError:
        print("Invalid integer parameter. Please provide a valid integer.")
