<?php
    include 'koneksi.php';
    require 'vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', 'No');
    $sheet->setCellValue('B1', 'Jenis Pendaftaran');
    $sheet->setCellValue('C1', 'Tanggal Masuk Sekolah');
    $sheet->setCellValue('D1', 'NIS');
    $sheet->setCellValue('E1', 'No Peserta Ujian');
    $sheet->setCellValue('F1', 'Pernah Paud');
    $sheet->setCellValue('G1', 'Pernah Tk');
    $sheet->setCellValue('H1', 'No Skhun');
    $sheet->setCellValue('I1', 'No Ijazah');
    $sheet->setCellValue('J1', 'Hobi');
    $sheet->setCellValue('K1', 'Cita - Cita');
    $sheet->setCellValue('L1', 'Nama Lengkap');
    $sheet->setCellValue('M1', 'Jenis Kelamin');
    $sheet->setCellValue('N1', 'NISN');
    $sheet->setCellValue('O1', 'NIK');
    $sheet->setCellValue('P1', 'Tempat Lahir');
    $sheet->setCellValue('Q1', 'Tanggal Lahir');
    $sheet->setCellValue('R1', 'Agama');
    $sheet->setCellValue('S1', 'Berkebutuhan Khusus');
    $sheet->setCellValue('T1', 'Alamat');
    $sheet->setCellValue('U1', 'RT');
    $sheet->setCellValue('V1', 'RW');
    $sheet->setCellValue('W1', 'Dusun');
    $sheet->setCellValue('X1', 'Kelurahan');
    $sheet->setCellValue('Y1', 'Kecamatan');
    $sheet->setCellValue('Z1', 'Kode Pos');
    $sheet->setCellValue('AA1', 'Tempat Tinggal');
    $sheet->setCellValue('AB1', 'Transportasi');
    $sheet->setCellValue('AC1', 'No Hp');
    $sheet->setCellValue('AD1', 'No Tlp');
    $sheet->setCellValue('AE1', 'Email');
    $sheet->setCellValue('AF1', 'Penerima Kps');
    $sheet->setCellValue('AG1', 'No Kps');
    $sheet->setCellValue('AH1', 'Kewarganegaraan');
    $sheet->setCellValue('AI1', 'Negara');
    $sheet->setCellValue('AJ1', 'Nama Ayah Kandung');
    $sheet->setCellValue('AK1', 'Tahun Lahir');
    $sheet->setCellValue('AL1', 'Pendidikan');
    $sheet->setCellValue('AM1', 'Pekerjaan');
    $sheet->setCellValue('AN1', 'Penghasilan Bulanan');
    $sheet->setCellValue('AO1', 'Berkebutuhan Khusus');
    $sheet->setCellValue('AP1', 'Nama Ibu Kandung');
    $sheet->setCellValue('AQ1', 'Tahun Lahir');
    $sheet->setCellValue('AR1', 'Pendidikan');
    $sheet->setCellValue('AS1', 'Pekerjaan');
    $sheet->setCellValue('AT1', 'Penghasilan');
    $sheet->setCellValue('AU1', 'Berkebutuhan Khusus');


    $host = "localhost";
    $user = "root";
    $password = "";
    $database = "praktik_formulir";
    $koneksi = mysqli_connect($host, $user, $password, $database);
    $sql = mysqli_query($koneksi, "SELECT * FROM peserta, datapribadi, dataayahkandung, dataibukandung");
    $i = 2;
    $no = 1;
    while ($row = mysqli_fetch_array($sql)) {
        $sheet->setCellValue('A'.$i, $no++);
        $sheet->setCellValue('B'.$i, $row['jenpend']);
        $sheet->setCellValue('C'.$i, $row['tglmsksklh']);
        $sheet->setCellValue('D'.$i, $row['nis']);
        $sheet->setCellValue('E'.$i, $row['nopeujian']);
        $sheet->setCellValue('F'.$i, $row['paud']);
        $sheet->setCellValue('G'.$i, $row['tk']);
        $sheet->setCellValue('H'.$i, $row['noserskhun']);
        $sheet->setCellValue('I'.$i, $row['noserijazah']);
        $sheet->setCellValue('J'.$i, $row['hobi']);
        $sheet->setCellValue('K'.$i, $row['cita']);
        $sheet->setCellValue('L'.$i, $row['namleng']);
        $sheet->setCellValue('M'.$i, $row['jkel']);
        $sheet->setCellValue('N'.$i, $row['nisn']);
        $sheet->setCellValue('O'.$i, $row['nik']);
        $sheet->setCellValue('P'.$i, $row['temlahir']);
        $sheet->setCellValue('Q'.$i, $row['tglahir']);
        $sheet->setCellValue('R'.$i, $row['agama']);
        $sheet->setCellValue('S'.$i, $row['kebukhusus']);
        $sheet->setCellValue('T'.$i, $row['alamat']);
        $sheet->setCellValue('U'.$i, $row['rt']);
        $sheet->setCellValue('V'.$i, $row['rw']);
        $sheet->setCellValue('W'.$i, $row['namdus']);
        $sheet->setCellValue('X'.$i, $row['namkel']);
        $sheet->setCellValue('Y'.$i, $row['kec']);
        $sheet->setCellValue('Z'.$i, $row['kodepos']);
        $sheet->setCellValue('AA'.$i, $row['ttinggal']);
        $sheet->setCellValue('AB'.$i, $row['transport']);
        $sheet->setCellValue('AC'.$i, $row['nohp']);
        $sheet->setCellValue('AD'.$i, $row['notelp']);
        $sheet->setCellValue('AE'.$i, $row['email']);
        $sheet->setCellValue('AF'.$i, $row['kpspkh']);
        $sheet->setCellValue('AG'.$i, $row['nokpspkh']);
        $sheet->setCellValue('AH'.$i, $row['kwn']);
        $sheet->setCellValue('AI'.$i, $row['namaneg']);
        $sheet->setCellValue('AJ'.$i, $row['namaayah']);
        $sheet->setCellValue('AK'.$i, $row['tlayah']);
        $sheet->setCellValue('AL'.$i, $row['pendayah']);
        $sheet->setCellValue('AM'.$i, $row['kerjayah']);
        $sheet->setCellValue('AN'.$i, $row['gajiayah']);
        $sheet->setCellValue('AO'.$i, $row['kebuayah']);
        $sheet->setCellValue('AP'.$i, $row['namaibu']);
        $sheet->setCellValue('AQ'.$i, $row['tlibu']);
        $sheet->setCellValue('AR'.$i, $row['pendibu']);
        $sheet->setCellValue('AS'.$i, $row['kerjaibu']);
        $sheet->setCellValue('AT'.$i, $row['gajibu']);
        $sheet->setCellValue('AU'.$i, $row['kebuibu']);
        $i++;
    }
    $styleArray = [
        'borders'=>[
            'allBorders'=>[
                'borderStyle'=> \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];

    $i = $i - 1;
    $sheet->getStyle('A1:AU1'.$i)->applyFromArray($styleArray);
    $writer = new Xlsx($spreadsheet);
    $writer->save('Report Data Formulir.xlsx');
?>