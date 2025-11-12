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


// ==================== SỰ KIỆN THAY ĐỔI SỐ BẢN GHI MỖI TRANG ====================
$(document).on('change', '#pageSizeSelect', function () {
    pageSize = parseInt($(this).val());
    currentPage = 1;
    filterData();
});

// ==================== SỰ KIỆN PHÂN TRANG ====================
$(document).on('click', '.page-link', function (e) {
    e.preventDefault();
    const page = $(this).data('page');
    if (page >= 1 && page <= totalPages && page !== currentPage) {
        currentPage = page;
        filterData(true);
    }
});
$(document).on('click', '#btnFilter', function (e) {
    e.preventDefault();
    currentPage = 1;
    isInitialLoad = true;
    filterData();
});

let tenNVDN = "";

$.getJSON("dist/data/json/Dm_NhanVien.json", data => {
    const nv = data.find(n => n.id === _idNVDN || n.ID === _idNVDN || n.Id === _idNVDN);
    if (nv) {
        tenNVDN = nv.ten || nv.Ten || nv.TenNhanVien || "";
        console.log("Tên nhân viên:", tenNVDN);
    } else {
        console.warn("Không tìm thấy nhân viên có ID =", idNVDN);
    }
});

// ==================== LỌC DỮ LIỆU ====================
let firstLoad = true;
function filterData(isPagination = false) {
    let tuNgay = $('#ngayTuNgay').val();
    let denNgay = $('#ngayDenNgay').val();
    let idKhoa = $('.tomselect-khoa').val() || 0;
    let idPhongBuong = $('.tomselect-phongBuong').val() || 0;
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
        idKhoa: idKhoa,
        idPhongBuong: idPhongBuong,
        page: currentPage,
        pageSize: pageSize
    }
    $.ajax({
        url: '/bao_cao_tiep_nhan/filter',
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
function ajaxFilterRequest(payload) {
    return new Promise((resolve, reject) => {
        $.ajax({
            url: '/bao_cao_tiep_nhan/filter',
            type: 'POST',
            data: payload,
            success: function (resp) { resolve(resp); },
            error: function (xhr, st, err) { reject(err || st || xhr); }
        });
    });
}

function fetchAllFilteredData(tuNgay, denNgay, idKhoa, idPhongBuong) {
    return new Promise((resolve, reject) => {
        const basePayload = {
            tuNgay: tuNgay || '',
            denNgay: denNgay || '',
            IdChiNhanh: _idcn || 0,
            idKhoa: idKhoa || 0,
            idPhongBuong: idPhongBuong || 0,
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
                    idKhoa: idKhoa || 0,
                    idPhongBuong: idPhongBuong || 0,
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

    if (!tuNgay && !denNgay) {
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
function doExportExcel(finalData, btn, originalHtml) {
    const requestData = {
        data: finalData,
        fromDate: $('#ngayTuNgay').val(),
        toDate: $('#ngayDenNgay').val(),
        idKhoa: $('.tomselect-khoa').val() || 0,
        idPhongBuong: $('.tomselect-phongBuong').val() || 0,
        doanhNghiep: window.doanhNghiep || null,
    };

    $.ajax({
        url: '/bao_cao_tiep_nhan/export/excel',
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
            a.download = `BaoCaoTiepNhan_${requestData.fromDate || 'all'}_den_${requestData.toDate || 'now'}.xlsx`;
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
    const idKhoa = $('.tomselect-httt').val() || 0;
    const idPhongBuong = $('.tomselect-nhanVien').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idKhoa, idPhongBuong)
            .then(allData => {
                window.filteredData = allData;
                doExportExcel(allData, btn, originalHtml);
            })
            .catch(err => {
                btn.html(originalHtml);
                btn.prop('disabled', false);
            });
    } else {
        doExportExcel(window.filteredData, btn, originalHtml);
    }
});

// ==================== XUẤT PDF ====================
function doExportPdf(finalData, btnElem) {
    const requestData = {
        data: finalData,
        fromDate: $('#ngayTuNgay').val(),
        toDate: $('#ngayDenNgay').val(),
        idKhoa: $('.tomselect-httt').val() || 0,
        idPhongBuong: $('.tomselect-nhanVien').val() || 0,
        doanhNghiep: window.doanhNghiep || null,
    };

    fetch("/bao_cao_tiep_nhan/export/pdf", {
        method: "POST",
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/pdf' },
        body: JSON.stringify(requestData)
    })
        .then(res => {
            if (!res.ok) throw new Error('Network response was not ok');
            return res.blob();
        })
        .then(blob => {
            const pdfUrl = URL.createObjectURL(blob);

            // Tạo iframe ẩn để mở file PDF
            const iframe = document.createElement('iframe');
            iframe.style.display = 'none';
            iframe.src = pdfUrl;
            document.body.appendChild(iframe);

            iframe.onload = function () {
                const printWindow = iframe.contentWindow;
                printWindow.focus();
                printWindow.print();
            };
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
    const idKhoa = $('.tomselect-httt').val() || 0;
    const idPhongBuong = $('.tomselect-nhanVien').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idKhoa, idPhongBuong)
            .then(allData => {
                window.filteredData = allData;
                doExportPdf(allData, btn);
            })
            .catch(err => {
                btn.innerHTML = '<i class="bi bi-file-earmark-pdf"></i> Xuất PDF';
                btn.disabled = false;
            });
    } else {
        doExportPdf(window.filteredData, btn);
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
    if (value == null) return "0";
    return Number(value).toLocaleString('en-US', {
        minimumFractionDigits: 0,
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
        // Gom nhóm theo nhân viên
        const groupedData = {};
        data.forEach(item => {
            const id = item.idKhoa || 0;
            const ten = item.tenKhoa || "Không rõ";
            const key = `${id}-${ten}`;

            if (!groupedData[key]) {
                groupedData[key] = {
                    idKhoa: id,
                    tenkhoa: ten,
                    items: []
                };
            }
            groupedData[key].items.push(item);
        });

        Object.keys(groupedData).forEach(key => {
            const group = groupedData[key];
            const items = group.items;
            console.log(group);
            // Header nhân viên
            let tongLuot = 0;
            let tongNam = 0;
            let tongNu = 0;
            let tongBHYT = 0;
            let tongKhongBHYT = 0;

            items.forEach(item => {
                tongLuot += Number(item.soLuotTiepNhan || 0);
                tongNam += Number(item.soLuongNam || 0);
                tongNu += Number(item.soLuongNu || 0);
                tongBHYT += Number(item.coBHYT || 0);
                tongKhongBHYT += Number(item.khongBHYT || 0);

            });
            const nvRow = `
            <tr class="fw-bold">
                <td colspan="6" class = "fw-bold text-start">${group.tenkhoa}</td>
                <td></td>
            </tr>`;
            tbody.append(nvRow);

            items.forEach((item, index) => {
                const stt = (currentPage - 1) * pageSize + index + 1;

                const row = `
                <tr>
                    <td class="text-nowrap text-center">${stt}</td>
                    <td class="text-nowrap text-center">${item.tenPhongBenh || ''}</td>
                    <td class="text-end text-nowrap">${formatCurrency(item.soLuotTiepNhan)}</td>
                    <td class="text-end text-nowrap">${formatCurrency(item.soLuongNam)}</td>
                    <td class="text-end text-nowrap">${formatCurrency(item.soLuongNu)}</td>
                    <td class="text-end text-nowrap">${formatCurrency(item.coBHYT)}</td>
                    <td class="text-end text-nowrap">${formatCurrency(item.khongBHYT)}</td>
                </tr>`;
                tbody.append(row);
            });

            // Tổng riêng cho nhân viên
            const totalRow = `
            <tr class="fw-bold table-secondary">
                <td colspan = "2" class="text-end text-nowrap"></td>
                <td class="text-end text-nowrap khongapdung">${formatCurrency(tongLuot)}</td>
                <td class="text-end text-nowrap">${formatCurrency(tongNam)}</td>
                <td class="text-end text-nowrap">${formatCurrency(tongNu)}</td>
                <td class="text-end text-nowrap">${formatCurrency(tongBHYT)}</td>
                <td class="text-end text-nowrap">${formatCurrency(tongKhongBHYT)}</td>
            </tr>`;
            tbody.append(totalRow);
        });
        const total = `
            <tr class="fw-bold table-secondary">
                <td colspan = "2" class="text-end text-nowrap"></td>
                <td class="text-end text-nowrap khongapdung">${formatCurrency(response.tongTiepNhan)}</td>
                <td class="text-end text-nowrap">${formatCurrency(response.tongNam)}</td>
                <td class="text-end text-nowrap">${formatCurrency(response.tongNu)}</td>
                <td class="text-end text-nowrap">${formatCurrency(response.tongCoBHYT)}</td>
                <td class="text-end text-nowrap">${formatCurrency(response.tongKhongBHYT)}</td>
            </tr>`;
        tbody.append(total);
    } else {
        tbody.append('<tr><td colspan="7" class="text-center">Không có dữ liệu</td></tr>');
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

$.getJSON("dist/data/json/DM_Khoa.json", dataKhoa => {
    listKhoa = dataKhoa
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
            className: ".tomselect-khoa",
            dieuKien: function (response) {
                return response.filter(x => x.id);
            }
        }
    ];

    configCb(configs, listKhoa);
});

function loadPhongBuong() {
    $.getJSON("dist/data/json/DM_PhongBuong.json", dataHTTT => {
        const idKhoa = parseInt($('.tomselect-khoa').val() || 0, 10);
        listHTTT = dataHTTT
            .filter(n =>
                (n.active === true || n.active === 1) &&
                (idKhoa === 0 || n.idKhoa === idKhoa)
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
                className: ".tomselect-phongBuong",
                dieuKien: function (response) {
                    return response.filter(x => x.id);
                }
            }
        ];

        configCb(configs, listHTTT);
    });
}

loadPhongBuong();

// bắt sự kiện thay đổi nhóm hàng
$(document).on('change', '.tomselect-khoa', function () {
    const ts = document.querySelector('.tomselect-phongBuong')?.tomselect;
    if (ts) {
        ts.clear();
        ts.clearOptions();
    }
    loadPhongBuong();
});