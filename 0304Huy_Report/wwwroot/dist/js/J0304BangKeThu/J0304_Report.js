// ==================== COMBOBOX ====================
function initAutocomplete(config) {
    const {
        inputId,
        dropdownId,
        hiddenIdId,
        data,
        getName,
        getId,
        getAbbr = () => "", // Optional abbreviation getter
        filterPredicate
    } = config;

    const input = document.getElementById(inputId);
    const dropdown = document.getElementById(dropdownId);
    const hiddenId = document.getElementById(hiddenIdId);
    let isMouseDownOnDropdown = false;
    let highlightedIndex = -1;
    let currentOptions = [];

    hiddenId.value = 0;

    function renderOptions(filter = "") {
        dropdown.innerHTML = "";
        highlightedIndex = 0;
        const normalizedFilter = removeAccents(filter.toLowerCase());

        currentOptions = data.filter(item => filterPredicate(item, normalizedFilter));

        currentOptions.forEach((item, index) => {
            const option = document.createElement('div');
            option.classList.add('option-item');

            const nameSpan = document.createElement('span');
            nameSpan.innerHTML = highlightMatch(getName(item), filter);
            nameSpan.style.flex = "1";
            option.appendChild(nameSpan);

            const abbr = getAbbr(item);
            if (abbr) {
                const abbrSpan = document.createElement('span');
                abbrSpan.innerHTML = highlightMatch(abbr, filter);
                abbrSpan.style.marginLeft = "10px";
                abbrSpan.style.color = "#888";
                abbrSpan.style.fontSize = "12px";
                option.appendChild(abbrSpan);
            }

            if (index === highlightedIndex) option.classList.add('highlight');

            option.addEventListener('mousedown', (e) => {
                e.preventDefault();
                selectOption(index);
            });

            dropdown.appendChild(option);
        });

        dropdown.style.display = currentOptions.length ? "block" : "none";
    }

    function updateHighlight() {
        const options = dropdown.querySelectorAll('.option-item');
        options.forEach((opt, idx) => {
            opt.classList.toggle('highlight', idx === highlightedIndex);
        });
    }

    function selectOption(index) {
        if (index >= 0 && index < currentOptions.length) {
            input.value = getName(currentOptions[index]);
            hiddenId.value = getId(currentOptions[index]);
            dropdown.style.display = "none";
        }
    }

    input.addEventListener('input', () => {
        if (input.value.trim() === "") {
            hiddenId.value = 0;
            dropdown.style.display = "none";
        } else {
            hiddenId.value = "";
            renderOptions(input.value);
        }
    });

    dropdown.addEventListener('mousedown', () => {
        isMouseDownOnDropdown = true;
    });

    input.addEventListener('blur', () => {
        setTimeout(() => {
            if (!isMouseDownOnDropdown) {
                if (hiddenId.value === "" && input.value.trim() !== "") {
                    input.value = "";
                    hiddenId.value = 0;
                }
            }
            isMouseDownOnDropdown = false;
            dropdown.style.display = "none";
        }, 100);
    });

    input.addEventListener('focus', () => renderOptions());

    input.addEventListener('input', () => {
        renderOptions(input.value);
    });

    window.addEventListener('load', () => {
        if (hiddenId.value && !input.value) {
            const selected = data.find(x => getId(x) == hiddenId.value);
            if (selected) {
                input.value = getName(selected);
            }
        }
    });

    input.addEventListener('keydown', (e) => {
        if (dropdown.style.display === "block") {
            if (e.key === "ArrowDown") {
                e.preventDefault();
                highlightedIndex = (highlightedIndex + 1) % currentOptions.length;
                updateHighlight();
            } else if (e.key === "ArrowUp") {
                e.preventDefault();
                highlightedIndex = (highlightedIndex - 1 + currentOptions.length) % currentOptions.length;
                updateHighlight();
            } else if (e.key === "Enter") {
                e.preventDefault();
                selectOption(highlightedIndex);
            }
        }
    });

    document.addEventListener('click', (e) => {
        const isClickInsideCombo = e.target.closest(`#${inputId}`) || e.target.closest(`#${dropdownId}`);
        if (!isClickInsideCombo) {
            if (hiddenId.value === "" && input.value.trim() !== "") {
                input.value = "";
                hiddenId.value = 0;
            }
            dropdown.style.display = "none";
        }
    });
}
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

// ==================== LỌC DỮ LIỆU ====================
let firstLoad = true;
function filterData(isPagination = false) {
    let tuNgay = $('#ngayTuNgay').val();
    let denNgay = $('#ngayDenNgay').val();
    let idHTTT = $('#IDHTTT').val() || 0;
    let idNhanVien = $('#IDNhanVien').val() || 0;
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
        page: currentPage,
        pageSize: pageSize
    }
    $.ajax({
        url: '/bang_ke_thu_ngoai_tru/filter',
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
            url: '/bang_ke_thu_ngoai_tru/filter',
            type: 'POST',
            data: payload,
            success: function (resp) { resolve(resp); },
            error: function (xhr, st, err) { reject(err || st || xhr); }
        });
    });
}

function fetchAllFilteredData(tuNgay, denNgay, idHTTT, idNhanVien) {
    return new Promise((resolve, reject) => {
        const basePayload = {
            tuNgay: tuNgay || '',
            denNgay: denNgay || '',
            IdChiNhanh: _idcn || 0,
            idHTTT: idHTTT || 0,
            idNhanVien: idNhanVien,
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
                    idNhanVien: idNhanVien,
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
    const idHTTT = $('#IDHTTT').val() || 0;
    const idNhanVien = $('#IDNhanVien').val() || 0;

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
function doExportExcel(finalData, btn, originalHtml) {
    const requestData = {
        data: finalData,
        fromDate: $('#ngayTuNgay').val(),
        toDate: $('#ngayDenNgay').val(),
        idHTTT: $('#IDHTTT').val() || 0,
        idNhanVien: $('#IDNhanVien').val() || 0,
        doanhNghiep: window.doanhNghiep || null
    };

    $.ajax({
        url: '/bang_ke_thu_ngoai_tru/export/excel',
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
    const idHTTT = $('#IDHTTT').val() || 0;
    const idNhanVien = $('#IDNhanVien').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idHTTT, idNhanVien)
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
        idHTTT: $('#IDHTTT').val() || 0,
        idNhanVien: $('#IDNhanVien').val() || 0,
        doanhNghiep: window.doanhNghiep || null
    };

    fetch("/bang_ke_thu_ngoai_tru/export/pdf", {
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
            a.download = `BangKeThuNgoaiTru_${requestData.fromDate || 'all'}_den_${requestData.toDate || 'now'}.pdf`;
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
    const idHTTT = $('#IDHTTT').val() || 0;
    const idNhanVien = $('#IDNhanVien').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idHTTT, idNhanVien)
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

    if (data.length > 0) { // Cần chỉnh lại chỗ này
        let tongHuy = 0;
        let tongHoan = 0;
        let tongSoTien = 0;
        data.forEach((item, index) => {
            tongHuy += Number(item.huy || item.Huy || 0);
            tongHoan += Number(item.hoan || item.Hoan || 0);
            tongSoTien += Number(item.soTien || item.SoTien || 0);
            const stt = (currentPage - 1) * pageSize + index + 1;
            const row = `
                <tr>
                    <td class="text-nowrap text-center">${stt}</td>
                    <td class="text-nowrap text-center">${item.maYTe || item.MaYTe || ''}</td>
                    <td class="text-nowrap text-start">${item.HoVaTen || item.hoVaTen || 'Không rõ'}</td>
                    <td class="text-nowrap text-center">${item.quyenSo || item.QuyenSo || 'Không rõ'}</td>
                    <td class="text-nowrap text-center">${item.soBienLai || item.SoBienLai || 'Không rõ'}</td>
                    <td class="text-nowrap text-center">${item.loai || item.Loai || 'Không rõ'}</td>
                    <td class="text-nowrap text-center">${formatDate(item.ngayThu || item.NgayThu)}</td>
                    <td class="text-nowrap text-end">${formatCurrency(item.huy || item.Huy)}</td>
                    <td class="text-nowrap text-end">${formatCurrency(item.hoan || item.Hoan)}</td>
                    <td class="text-nowrap text-end">${formatCurrency(item.soTien || item.SoTien)}</td>
                </tr>
            `;
            tbody.append(row);
        });
        const totalRow = `
                <tr class="fw-bold">
                    <td colspan="7" class="text-center text-nowrap">Tổng cộng</td>
                    <td class="text-end text-nowrap">${formatCurrency(tongHuy) }</td>
                    <td class="text-end text-nowrap">${formatCurrency(tongHoan)}</td>
                    <td class="text-end text-nowrap">${formatCurrency(tongSoTien)}</td>
                </tr>
            `;
        tbody.append(totalRow);
    } else {
        tbody.append('<tr><td colspan="10" class="text-center">Không có dữ liệu</td></tr>');
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
document.addEventListener("DOMContentLoaded", () => {

    const dataHTTT = convertData(
        provincesDataHTTT,
        item => item.ten || '',
        item => item.id 
    );

    const dataNhanVien = convertData(
        provincesDataNhanVien,
        item => item.TenNhanVien || '', 
        item => item.ID
    );

    initAutocomplete({
        inputId: 'comboBox',
        dropdownId: 'dropdownList',
        hiddenIdId: 'IDHTTT',
        data: dataHTTT,
        getName: item => item.ten || '',
        getId: item => item.id,
        getAbbr: item => item.viettat || '',
        filterPredicate: (item, normalizedFilter) =>
            removeAccents((item.ten || '').toLowerCase()).includes(normalizedFilter) ||
            removeAccents((item.viettat || '').toLowerCase()).startsWith(normalizedFilter)
    });

    initAutocomplete({
        inputId: 'comboBox2',
        dropdownId: 'dropdownList2',
        hiddenIdId: 'IDNhanVien',
        data: dataNhanVien,
        getName: item => item.ten || '',
        getId: item => item.id,
        getAbbr: item => item.viettat || '',
        filterPredicate: (item, normalizedFilter) =>
            removeAccents((item.ten || '').toLowerCase()).includes(normalizedFilter) ||
            removeAccents((item.viettat || '').toLowerCase()).startsWith(normalizedFilter)
    });
});
