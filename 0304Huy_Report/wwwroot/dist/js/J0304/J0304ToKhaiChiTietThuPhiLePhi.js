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
    let idHTTT = $('.tomselect-httt').val() || 0;
    let idNhanVien = $('.tomselect-nhanVien').val() || 0;
    let idLoai = $('.tomselect-loai').val() || 0;
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
        idHTTT: idHTTT,
        idNhanVien: idNhanVien,
        idLoai: idLoai,
        page: currentPage,
        pageSize: pageSize
    }
    $.ajax({
        url: '/to_khai_chi_tiet_thu_phi_le_phi/filter',
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
            url: '/to_khai_chi_tiet_thu_phi_le_phi/filter',
            type: 'POST',
            data: payload,
            success: function (resp) { resolve(resp); },
            error: function (xhr, st, err) { reject(err || st || xhr); }
        });
    });
}

function fetchAllFilteredData(tuNgay, denNgay, idNhanVien) {
    return new Promise((resolve, reject) => {
        const basePayload = {
            tuNgay: tuNgay || '',
            denNgay: denNgay || '',
            IdChiNhanh: _idcn || 0,
            idHTTT: idHTTT || 0,
            idNhanVien: idNhanVien || 0,
            idLoai: idLoai || 0,
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
                    idHTTT: idHTTT || 0,
                    idNhanVien: idNhanVien || 0,
                    idLoai: idLoai || 0,
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
        idHTTT: $('.tomselect-httt').val() || 0,
        idNhanVien: $('.tomselect-nhanVien').val() || 0,
        idLoai: $('.tomselect-loai').val() || 0,
        doanhNghiep: window.doanhNghiep || null,
        TenNVDN: tenNVDN,
    };

    $.ajax({
        url: '/to_khai_chi_tiet_thu_phi_le_phi/export/excel',
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
            a.download = `ToKhaiChiTietThuPhiLePhi_${requestData.fromDate || 'all'}_den_${requestData.toDate || 'now'}.xlsx`;
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
    const idHTTT = $('.tomselect-httt').val() || 0;
    const idNhanVien = $('.tomselect-nhanVien').val() || 0;
    const idLoai = $('.tomselect-loai').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idHTTT, idNhanVien, idLoai)
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
        idHTTT: $('.tomselect-httt').val() || 0,
        idNhanVien: $('.tomselect-nhanVien').val() || 0,
        idLoai: $('.tomselect-loai').val() || 0,
        doanhNghiep: window.doanhNghiep || null,
        TenNVDN: tenNVDN,
    };

    fetch("/to_khai_chi_tiet_thu_phi_le_phi/export/pdf", {
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
    const idHTTT = $('.tomselect-httt').val() || 0;
    const idNhanVien = $('.tomselect-nhanVien').val() || 0;
    const idLoai = $('.tomselect-loai').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idHTTT, idNhanVien, idLoai)
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

    if (data && data.length > 0) {

        const groupedData = data.reduce((acc, item) => {
            const nhanVienKey = `${item.idNhanVien || item.IDNhanVien || 0}|${item.tenNhanVien || item.TenNhanVien || 'Không rõ nhân viên'}`;
            const loaiHDKey = item.loaiHoaDon || item.LoaiHoaDon || 'Không rõ loại hóa đơn';

            if (!acc[nhanVienKey]) acc[nhanVienKey] = {};
            if (!acc[nhanVienKey][loaiHDKey]) acc[nhanVienKey][loaiHDKey] = [];
            acc[nhanVienKey][loaiHDKey].push(item);

            return acc;
        }, {});

        Object.keys(groupedData).forEach(nvKey => {
            const [idNhanVien, tenNhanVien] = nvKey.split('|');
            tbody.append(`
            <tr>
                <td colspan="8" class="fw-bold text-start bg-light">
                    ${tenNhanVien}
                </td>
            </tr>
        `);

            const loaiHDDict = groupedData[nvKey];
            Object.keys(loaiHDDict).forEach(loaiHD => {
                tbody.append(`
                <tr>
                    <td colspan="8" class="fw-bold text-start ps-3">
                        Loại HĐ: ${loaiHD}
                    </td>
                </tr>
            `);
                let stt = 1;
                loaiHDDict[loaiHD].forEach(item => {
                    const row = `
                    <tr>
                        <td class="text-center text-nowrap fw-bold" style="width: 65px;">${stt++}</td>
                        <td class="text-start text-nowrap">${item.seriVaQuyen || item.SeriVaQuyen || ''}</td>
                        <td class="text-center text-nowrap">${item.soLan_soBLHDthu || item.SoLan_soBLHDthu || ''}</td>
                        <td class="text-center text-nowrap fw-bold">${item.soLuongHDSuDung || item.SoLuongHDSuDung || ''}</td>
                        <td class="text-center text-nowrap">${formatCurrency(item.tongSoTien || item.TongSoTien || 0)}</td>
                        <td class="text-center text-nowrap">${formatCurrency(item.huy_hoan || item.Huy_Hoan || 0)}</td>
                        <td class="text-end text-nowrap">
                          ${(() => {
                            const val = item.soTienThucThu || item.SoTienThucThu || 0;
                            return val < 0
                                ? '(' + formatCurrency(Math.abs(val)) + ')'
                                : formatCurrency(val);
                        })()}
                        </td>
                        <td class="text-start text-nowrap fst-italic">${item.ghiChu || item.GhiChu || ''}</td>
                    </tr>
                `;
                    tbody.append(row);
                });
            });
        });

        // ===== DÒNG TỔNG CỘNG =====
        const totalRow = `
        <tr class="fw-bold">
            <td colspan="4" class="text-end text-nowrap">Tổng cộng:</td>
            <td class="text-end text-nowrap">${formatCurrency(response.allSoTien || response.AllSoTien)}</td>
            <td class="text-end text-nowrap">${formatCurrency(response.allHoan_Huy || response.AllHoan_Huy)}</td>
            <td class="text-end text-nowrap">${formatCurrency(response.allTienThucThu || response.AllTienThucThu)}</td>
            <td></td>
        </tr>
    `;
        tbody.append(totalRow);
    } else {
        tbody.append('<tr><td colspan="8" class="text-center">Không có dữ liệu</td></tr>');
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

$.getJSON("dist/data/json/Dm_NhanVien.json", dataNhanVien => {
    listNhanVien = dataNhanVien
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
            className: ".tomselect-nhanVien",
            dieuKien: function (response) {
                return response.filter(x => x.id);
            }
        }
    ];

    configCb(configs, listNhanVien);
});

$.getJSON("dist/data/json/DM_HTTT.json", dataHTTT => {
    listHTTT = dataHTTT
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
            className: ".tomselect-httt",
            dieuKien: function (response) {
                return response.filter(x => x.id);
            }
        }
    ];

    configCb(configs, listHTTT);
});

// ============== TOM SELECT ======================
function configCb2(configs, dataSource) {
    configs.forEach(cfg => {
        let result = cfg.dieuKien ? cfg.dieuKien(dataSource) : dataSource;

        // Lấy element
        let el = document.querySelector(cfg.className);
        if (!el) return;

        // Nếu đã có TomSelect rồi thì clear + add lại options
        if (el.tomselect) {
            el.tomselect.clearOptions();
            el.tomselect.addOptions(result);
            el.tomselect.refreshOptions(false);
        } else {
            // Nếu chưa có thì init mới
            new TomSelect(cfg.className, {
                options: result,
                valueField: "id",
                labelField: "ten",
                searchField: ["ten", "alias"],
                placeholder: cfg.placeholder,
                maxItems: 1,
                render: {
                    option: function (data, escape) {
                        return `
                            <div class="border-0" style="display:flex; justify-content:space-between; width:100%; border: none !important">
                                <span>${escape(data.ten)}</span>
                                <span style="color:gray; font-size:10px; margin-left:10px;">${escape(data.viettat || "")}</span>
                            </div>`;
                    },
                    item: function (data, escape) {
                        return `
                            <div class="border-0" style="display:flex; justify-content:space-between; width:100%; border: none !important;">
                                <span>${escape(data.ten)}</span>
                                <span style="color:gray; font-size:10px; margin-left:10px;">${escape(data.viettat || "")}</span>
                            </div>`;
                    }
                }
            });
        }
    });
}

(function () {
    const dataLoai = [
        { id: 1, ten: "DV kỹ thuật", viettat: "DVKT" },
        { id: 2, ten: "Thuốc", viettat: "T" }
    ];

    const listLoai = dataLoai.map(n => ({
        ...n,
        alias: n.viettat?.trim() !== ""
            ? n.viettat.toUpperCase()
            : n.ten.trim().split(/\s+/).map(w => w.charAt(0).toUpperCase()).join("")
    }));

    const configs = [
        {
            className: ".tomselect-loai",
            dieuKien: response => response.filter(x => x.id)
        }
    ];

    configCb(configs, listLoai);
})();