// ==================== BIẾN GLOBAL PHÂN TRANG ====================
let currentPage = 1;
let pageSize = 20;
let totalRecords = 0;
let totalPages = 0;
let isInitialLoad = true;

// ==================== RENDER PHÂN TRANG ====================
function renderPagination() {
    const pagination = $('#pagination');
    pagination.empty();

    const pages = Math.max(1, totalPages || Math.ceil(totalRecords / pageSize || 1));
    if (currentPage > pages) currentPage = pages;

    $('#pageInfo').text(`Trang ${currentPage}/${pages} - Tổng ${totalRecords} bản ghi`);

    pagination.append(`
        <li class="page-item ${currentPage === 1 ? 'disabled' : ''}">
            <a class="page-link" href="#" data-page="${Math.max(1, currentPage - 1)}">Trước</a>
        </li>
    `);

    const visibleCount = 3;
    let startPage = Math.max(1, currentPage - 1);
    let endPage = Math.min(pages, startPage + visibleCount - 1);

    if (endPage - startPage + 1 < visibleCount) {
        startPage = Math.max(1, endPage - visibleCount + 1);
    }

    for (let i = startPage; i <= endPage; i++) {
        pagination.append(`
            <li class="page-item ${i === currentPage ? 'active' : ''}">
                <a class="page-link" href="#" data-page="${i}">${i}</a>
            </li>
        `);
    }

    pagination.append(`
        <li class="page-item ${currentPage === pages ? 'disabled' : ''}">
            <a class="page-link" href="#" data-page="${Math.min(pages, currentPage + 1)}">Sau</a>
        </li>
    `);
}

let currentUrl = '/bao_cao_so_luong_nhap_xuat/filterNhap';
let currentUrlPDF = '/bao_cao_so_luong_nhap_xuat/exportnhap/pdf';
let currentUrlExcel = '/bao_cao_so_luong_nhap_xuat/exportnhap/excel';

// ==================== SỰ KIỆN PHÂN TRANG ====================
$(document).on('click', '#btnFilterNhap', function (e) {
    e.preventDefault();
    currentPage = 1;
    currentUrl = '/bao_cao_so_luong_nhap_xuat/filterNhap';
    currentUrlPDF = '/bao_cao_so_luong_nhap_xuat/exportnhap/pdf';
    currentUrlExcel = '/bao_cao_so_luong_nhap_xuat/exportnhap/excel';
    filterData(currentUrl);
});

$(document).on('click', '#btnFilterXuat', function (e) {
    e.preventDefault();
    currentPage = 1;
    currentUrl = '/bao_cao_so_luong_nhap_xuat/filterXuat';
    currentUrlPDF = '/bao_cao_so_luong_nhap_xuat/exportxuat/pdf';
    currentUrlExcel = '/bao_cao_so_luong_nhap_xuat/exportxuat/excel';
    filterData(currentUrl);
});

$(document).on('click', '.page-link', function (e) {
    e.preventDefault();
    const page = $(this).data('page');
    if (page >= 1 && page <= totalPages && page !== currentPage) {
        currentPage = page;
        filterData(currentUrl, true);
    }
});

// ==================== SỰ KIỆN THAY ĐỔI SỐ BẢN GHI MỖI TRANG ====================
$(document).on('change', '#pageSizeSelect', function () {
    pageSize = parseInt($(this).val());
    currentPage = 1;
    filterData(currentUrl);
});

// ==================== LỌC DỮ LIỆU ====================
let firstLoad = true;
function filterData(url, isPagination = false) {
    let tuNgay = $('#ngayTuNgay').val();
    let denNgay = $('#ngayDenNgay').val();
    let idKhoHang = 14;
    let idNhomHang = $('.tomselect-nhomHang').val() || 0;
    let idHangHoa = $('.tomselect-hangHoa').val() || 0;
    if (!isPagination) {
        firstLoad = true;
    }
    if (!isPagination && (!tuNgay || !denNgay)) {
        toastr.error("Vui lòng chọn từ ngày và đến ngày");
        return;
    }

    function parseDMY(s) {
        const p = s.split('-');
        return new Date(p[2], p[1] - 1, p[0]);
    }

    if (!isPagination && parseDMY(tuNgay) > parseDMY(denNgay)) {
        tuNgay = denNgay;
        $('#ngayTuNgay').val(tuNgay);
    }

    $('#loadingSpinner').show();
    $('.table-wrapper').css('opacity', '0.5');

    let payload = {
        tuNgay: tuNgay,
        denNgay: denNgay,
        IdChiNhanh: _idcn,
        idKhoHang: idKhoHang,
        idNhomHang: idNhomHang,
        idHangHoa: idHangHoa,
        page: currentPage,
        pageSize: pageSize
    }
    $.ajax({
        url: url,
        type: 'POST',
        data: payload,
        success: function (response) {
            console.log(response);
            if (response.success) {
                updateTable(response);
                window.filteredData = Array.isArray(response.data) ? response.data : (response.data ? [response.data] : []);
                totalRecords = response.totalRecords || totalRecords;
                totalPages = response.totalPages || totalPages;
                window.doanhNghiep = response.doanhNghiep || null;

                if (window.filteredData.length === 0) {
                    toastr.warning("Không có dữ liệu");
                } else if (firstLoad) {
                    toastr.success("Tải dữ liệu thành công");
                    firstLoad = false;
                }
            } else {
                toastr.error("Không có dữ liệu");
            }
        },
        complete: function () {
            $('#loadingSpinner').hide();
            $('.table-wrapper').css('opacity', '1');
        }
    });
}

// ==================== HÀM HỖ TRỢ LẤY TOÀN BỘ DỮ LIỆU ====================
function ajaxFilterRequest(url, payload) {
    return new Promise((resolve, reject) => {
        $.ajax({
            url: url,
            type: 'POST',
            data: payload,
            success: function (resp) { resolve(resp); },
            error: function (xhr, st, err) { reject(err || st || xhr); }
        });
    });
}

function fetchAllFilteredData(tuNgay, denNgay, idKhoHang, idNhomHang, idHangHoa) {
    return new Promise((resolve, reject) => {
        const basePayload = {
            tuNgay: tuNgay || '',
            denNgay: denNgay || '',
            IdChiNhanh: _idcn || 0,
            idKhoHang: idKhoHang,
            idNhomHang: idNhomHang ||0,
            idHangHoa: idHangHoa ||0,
            page: 1,
            pageSize: pageSize
        };

        ajaxFilterRequest(basePayload).then(firstResp => {
            if (!firstResp || !firstResp.success) {
                reject(firstResp || 'Lỗi khi lấy dữ liệu trang 1');
                return;
            }
            const firstData = Array.isArray(firstResp.data) ? firstResp.data : (firstResp.data ? [firstResp.data] : []);
            const tp = firstResp.totalPages || 1;

            if (tp <= 1) {
                resolve(firstData);
                return;
            }

            const promises = [];
            for (let p = 2; p <= tp; p++) {
                const payload = {
                    tuNgay: tuNgay || '',
                    denNgay: denNgay || '',
                    IdChiNhanh: _idcn,
                    idKhoHang: idKhoHang,
                    idNhomHang: idNhomHang || 0,
                    idHangHoa: idHangHoa || 0,
                    page: p,
                    pageSize: pageSize
                };
                promises.push(ajaxFilterRequest(payload));
            }

            Promise.all(promises)
                .then(results => {
                    const pagesData = results.map(r => Array.isArray(r.data) ? r.data : (r.data ? [r.data] : []));
                    const all = firstData.concat(...pagesData);
                    resolve(all);
                })
                .catch(err => {
                    reject(err);
                });
        }).catch(err => reject(err));
    });
}

// ==================== KIỂM TRA DỮ LIỆU XUẤT ====================
function validateExportDatesAndData() {
    const tuNgay = $('#ngayTuNgay').val();
    const denNgay = $('#ngayDenNgay').val();

    if (!tuNgay && !denNgay ) {
        if (!window.filteredData || window.filteredData.length === 0) {
            toastr.error("Không có dữ liệu để xuất");
            return false;
        }
        return true;
    }
    if ((tuNgay && !denNgay) || (!tuNgay && denNgay)) {
        toastr.error("Vui lòng chọn cả từ ngày và đến ngày");
        return false;
    }

    function parseDMY(s) {
        const parts = s.split('-');
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    if (parseDMY(tuNgay) > parseDMY(denNgay)) {
        toastr.error("Từ ngày phải nhỏ hơn hoặc bằng đến ngày");
        return false;
    }
    if (!window.filteredData || window.filteredData.length === 0) {
        toastr.error("Không có dữ liệu để xuất");
        return false;
    }
    return true;
}

// ==================== XUẤT EXCEL ====================
function doExportExcel(url, finalData, btn, originalHtml) {
    const requestData = {
        data: finalData,
        fromDate: $('#ngayTuNgay').val(),
        toDate: $('#ngayDenNgay').val(),
        idKhoHang: 14,
        idNhomHang: $('.tomselect-nhomHang').val() || 0,
        idHangHoa: $('.tomselect-hangHoa').val() || 0,
        doanhNghiep: window.doanhNghiep || null
    };

    $.ajax({
        url: url,
        type: 'POST',
        contentType: 'application/json',
        data: JSON.stringify(requestData),
        xhrFields: { responseType: 'blob' },
        success: function (data, status, xhr) {
            const contentType = xhr.getResponseHeader('content-type') || '';
            if (!contentType.includes('spreadsheet') && !contentType.includes('vnd.openxmlformats')) {
                return;
            }
            const blob = new Blob([data], { type: contentType });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `BangKeThuNgoaiTru_${requestData.fromDate || 'all'}_den_${requestData.toDate || 'now'}.xlsx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            toastr.success("Xuất Excel thành công");
        },
        error: function () {
            console.error("Error exporting Excel data");
            toastr.error("Xuất Excel thất bại");
        },
        complete: function () {
            btn.html(originalHtml);
            btn.prop('disabled', false);
        }
    });
}

$('#btnExportExcel').off('click').on('click', function (e) {
    e.preventDefault();
    if (!validateExportDatesAndData()) return;

    const btn = $(this);
    const originalHtml = btn.html();
    btn.html('<span class="spinner-border spinner-border-sm"></span> Đang tạo');
    btn.prop('disabled', true);

    const tu = $('#ngayTuNgay').val();
    const den = $('#ngayDenNgay').val();
    const idKhoHang = 14;
    const idNhomHang = $('.tomselect-nhomHang').val() || 0;
    const idHangHoa = $('.tomselect-hangHoa').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idKhoHang, idNhomHang, idHangHoa)
            .then(allData => {
                window.filteredData = allData;
                doExportExcel(currentUrlExcel, allData, btn, originalHtml);
            })
            .catch(err => {
                btn.html(originalHtml);
                btn.prop('disabled', false);
            });
    } else {
        doExportExcel(currentUrlExcel, window.filteredData, btn, originalHtml);
    }
});

// ==================== XUẤT PDF ====================
function doExportPdf(url, finalData, btnElem) {
    const requestData = {
        data: finalData,
        fromDate: $('#ngayTuNgay').val(),
        toDate: $('#ngayDenNgay').val(),
        idKhoHang: 14,
        idNhomHang: $('.tomselect-nhomHang').val() || 0,
        idHangHoa: $('.tomselect-hangHoa').val() || 0,
        doanhNghiep: window.doanhNghiep || null
    };

    fetch(url, {
        method: "POST",
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/pdf' },
        body: JSON.stringify(requestData)
    })
        .then(res => {
            if (!res.ok) throw new Error('Network response was not ok');
            return res.blob();
        })
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `BaoCaoSoLuong_${requestData.fromDate || 'all'}_den_${requestData.toDate || 'now'}.pdf`;
            a.click();
            window.URL.revokeObjectURL(url);
            toastr.success("Xuất PDF thành công");
        })
        .catch(error => {
            console.error('Error exporting PDF:', error);
            toastr.error("Xuất PDF thất bại");
        })
        .finally(() => {
            btnElem.innerHTML = '<i class="bi bi-file-earmark-pdf"></i> Xuất PDF';
            btnElem.disabled = false;
        });
}

$('#btnExportPDF').off('click').on('click', function (e) {
    e.preventDefault();
    if (!validateExportDatesAndData()) return;

    const btn = this;
    btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Đang tạo';
    btn.disabled = true;

    const tu = $('#ngayTuNgay').val();
    const den = $('#ngayDenNgay').val();
    const idKhoHang = 14;
    const idNhomHang = $('.tomselect-nhomHang').val() || 0;
    const idHangHoa = $('.tomselect-hangHoa').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idKhoHang, idNhomHang, idHangHoa)
            .then(allData => {
                window.filteredData = allData;
                doExportPdf(currentUrlPDF, allData, btn);
            })
            .catch(err => {
                btn.innerHTML = '<i class="bi bi-file-earmark-pdf"></i> Xuất PDF';
                btn.disabled = false;
            });
    } else {
        doExportPdf(currentUrlPDF, window.filteredData, btn);
    }
});


// ==================== ĐỊNH DẠNG NGÀY XUẤT RA BẢNG ====================
function formatDate(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    if (isNaN(date)) return dateString;
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
}

function formatCurrency(value) {
    return (value || 0.00).toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}


// ==================== CẬP NHẬT BẢNG ====================
function updateTable(response) {
    const tbody = $('.container_Team3.right tbody');
    tbody.empty();

    if (response.totalRecords !== undefined) {
        totalRecords = response.totalRecords;
        totalPages = response.totalPages;
        currentPage = response.currentPage || 1;
        $('#pageInfo').text(`Trang ${currentPage}/${totalPages} - Tổng ${totalRecords} bản ghi`);
        renderPagination();
    }

    let data = [];
    if (Array.isArray(response)) {
        data = response;
    } else if (response && response.data) {
        data = Array.isArray(response.data) ? response.data : [response.data];
    }

    if (data.length > 0) {
        // Gom nhóm theo công ty
        const groupedData = {};
        data.forEach(item => {
            const tenNhomHang = item.tenNhomHang || item.tenNhomHang || "Không rõ";
            if (!groupedData[tenNhomHang]) {
                groupedData[tenNhomHang] = [];
            }
            groupedData[tenNhomHang].push(item);
        });

        Object.keys(groupedData).forEach(tenNhomHang => {
            // Header công ty
            const companyRow = `
            <tr class="fw-bold">
                <td colspan="15" class="text-center tong-col">${tenNhomHang}</td>
            </tr>`;
            tbody.append(companyRow);

            // Reset tổng cho công ty này
            let tongDongGoi = 0;
            let tongLe = 0;
            let tongDonGiaDongGoi = 0;
            let tongDonGiaLe = 0;
            let tongThanhTien = 0;

            groupedData[tenNhomHang].forEach((item, index) => {
                //tongDongGoi += Number(item.slDongGoi || item.SLDongGoi || 0);
                //tongLe += Number(item.slLe || item.SLLe || 0);
                //tongDonGiaDongGoi += Number(item.donGiaDongGoi || item.DonGiaDongGoi || 0);
                //tongDonGiaLe += Number(item.donGiaLe || item.DonGiaLe || 0);
                //tongThanhTien += Number(item.thanhTien || item.ThanhTien || 0);
                const stt = (currentPage - 1) * pageSize + index + 1;

                const row = `
                <tr>
                    <td class="text-nowrap text-center">${stt}</td>
                    <td class="text-start" style="min-width: 450px; max-width: 500px;">${item.tenThuoc || item.TenThuoc || 'Không rõ'}</td>
                    <td class="text-nowrap text-end">${item.thang1 || item.Thang1 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang2 || item.Thang2 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang3 || item.Thang3 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang4 || item.Thang4 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang5 || item.Thang5 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang6 || item.Thang6 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang7 || item.Thang7 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang8 || item.Thang8 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang9 || item.Thang9 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang10 || item.Thang10 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang11 || item.Thang11 || 0}</td>
                    <td class="text-nowrap text-end">${item.thang12 || item.Thang12 || 0}</td>
                    <td class="text-nowrap text-end">${item.tongCong || item.tongCong || 0}</td>
                </tr>`;
                tbody.append(row);
            });

            // Tổng riêng cho công ty
            //const totalRow = `
            //<tr class="fw-bold table-secondary">
            //    <td colspan="3" class="text-end text-nowrap">Tổng cộng: </td>
            //    <td colspan="7" class="text-end text-nowrap khongapdung"></td>
            //    <td class="text-end text-nowrap khongapdung">${tongDongGoi}</td>
            //    <td class="text-end text-nowrap">${tongLe}</td>
            //    <td class="text-end text-nowrap">${formatCurrency(tongDonGiaDongGoi)}</td>
            //    <td class="text-end text-nowrap">${formatCurrency(tongDonGiaLe)}</td>
            //    <td class="text-end text-nowrap">${formatCurrency(tongThanhTien)}</td>
            //</tr>`;
            //tbody.append(totalRow);
        });
    } else {
        tbody.append('<tr><td colspan="15" class="text-center">Không có dữ liệu</td></tr>')
    }
}

// ==================== KHI TẢI TRANG ====================
$(document).ready(function () {
    const config = {
        tuNgay: 'ngayTuNgay', // id ô input
        denNgay: 'ngayDenNgay',
        tuNgayIcon: 'ngayTuNgay-icon', // id icon
        denNgayIcon: 'ngayDenNgay-icon',
        tuNgayDatepicker: 'ngayTuNgay-datepicker', // id cả cụm datepicker
        denNgayDatepicker: 'ngayDenNgay-datepicker'
    };

    initDateInputFormatting(config);
    initDatePicker(config);
    initDateRangeConstraint(config);
});

// ==================== LOAD COMBOBOX ====================
$.getJSON("dist/data/json/Dm_NhomHang.json", dataNhomHang => {
    listNhomHang = dataNhomHang
        .filter(n =>
            (n.active === true || n.active === 1)
        )
        .map(n => ({
            ...n,
            alias: n.viettat?.trim() !== ""
                ? n.viettat.toUpperCase()
                : n.ten.trim().split(/\s+/).map(w => w.charAt(0).toUpperCase()).join("")
        }));
    // config cho TomSelect
    const configs = [
        {
            className: ".tomselect-nhomHang",
            dieuKien: function (response) {
                return response.filter(x => x.id);
            }
        }
    ];

    configCb(configs, listNhomHang);
});

$.getJSON("dist/data/json/Dm_HangHoa.json", dataHangHoa => {
    listHangHoa = dataHangHoa
        .filter(n =>
            (n.active === true || n.active === 1)
        )
        .map(n => ({
            ...n,
            alias: n.viettat?.trim() !== ""
                ? n.viettat.toUpperCase()
                : n.ten.trim().split(/\s+/).map(w => w.charAt(0).toUpperCase()).join("")
        }));
    // config cho TomSelect
    const configs = [
        {
            className: ".tomselect-hangHoa",
            dieuKien: function (response) {
                return response.filter(x => x.id);
            }
        }
    ];

    configCb(configs, listHangHoa);
});