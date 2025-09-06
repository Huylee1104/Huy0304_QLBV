// ==================== ĐỊNH DẠNG NGÀY NHẬP ====================
function initDateInputFormatting() {
    const dateInputIds = ["ngayTuNgay", "ngayDenNgay"];

    dateInputIds.forEach(function (id) {
        const input = document.getElementById(id);
        if (!input) return;

        input.addEventListener("input", function () {
            let value = input.value.replace(/\D/g, "");
            let formatted = "";
            let selectionStart = input.selectionStart;

            if (value.length > 0) formatted += value.substring(0, 2);
            if (value.length >= 3) formatted += "-" + value.substring(2, 4);
            if (value.length >= 5) formatted += "-" + value.substring(4, 8);

            if (formatted !== input.value) {
                const prevLength = input.value.length;
                input.value = formatted;
                const newLength = formatted.length;
                const diff = newLength - prevLength;
                input.setSelectionRange(selectionStart + diff, selectionStart + diff);
            }
        });

        input.addEventListener("click", function () {
            const pos = input.selectionStart;
            if (pos <= 2) input.setSelectionRange(0, 2);
            else if (pos <= 5) input.setSelectionRange(3, 5);
            else input.setSelectionRange(6, 10);
        });

        input.addEventListener("keydown", function (e) {
            const pos = input.selectionStart;
            let val = input.value;

            if (e.key === "Backspace" && (pos === 3 || pos === 6)) {
                e.preventDefault();
                input.value = val.slice(0, pos - 1) + val.slice(pos);
                input.setSelectionRange(pos - 1, pos - 1);
            }
            if (e.key === "Delete" && (pos === 2 || pos === 5)) {
                e.preventDefault();
                input.value = val.slice(0, pos) + val.slice(pos + 1);
                input.setSelectionRange(pos, pos);
            }
        });
        $('#datepicker-icon').on('click', function () {
            $("#ngayTuNgay").datepicker('show');
        });
        $('#datepicker-icon2').on('click', function () {
            $("#ngayDenNgay").datepicker('show');
        });
    });
}

// ==================== DATEPICKER ====================
function initDatePicker() {
    $('[id="ngayTuNgay"], [id="ngayDenNgay"]').datepicker({
        format: 'dd-mm-yyyy',
        autoclose: true,
        language: 'vi',
        todayHighlight: true,
        orientation: 'bottom auto',
        weekStart: 1
    });
}

// ==================== COMBOBOX ====================
// Common helper functions
function removeAccents(str) {
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function highlightMatch(text, keyword) {
    if (!keyword) return text;

    const normalizedText = removeAccents(text).toLowerCase();
    const normalizedKeyword = removeAccents(keyword).toLowerCase();

    const startIndexNormalized = normalizedText.indexOf(normalizedKeyword);
    if (startIndexNormalized === -1) return text;

    let startIndexOriginal = 0;
    let count = 0;
    for (let i = 0; i < text.length; i++) {
        if (removeAccents(text[i]).toLowerCase() !== '') {
            if (count === startIndexNormalized) {
                startIndexOriginal = i;
                break;
            }
            count++;
        }
    }

    let endIndexOriginal = startIndexOriginal;
    let count2 = 0;
    for (let i = startIndexOriginal; i < text.length; i++) {
        if (removeAccents(text[i]).toLowerCase() !== '') {
            count2++;
        }
        if (count2 === normalizedKeyword.length) {
            endIndexOriginal = i + 1;
            break;
        }
    }

    return (
        text.substring(0, startIndexOriginal) +
        '<span class="highlight-text">' +
        text.substring(startIndexOriginal, endIndexOriginal) +
        '</span>' +
        text.substring(endIndexOriginal)
    );
}

// Factory function to initialize autocomplete
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
        idNhanVien: idNhanVien,
        page: currentPage,
        pageSize: pageSize
    }
    $.ajax({
        url: '/bang_ke_bien_lai_hoan_ung/filter',
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
            url: '/bang_ke_bien_lai_hoan_ung/filter',
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
    const idNhanVien = $('#IDNhanVien').val() || 0;

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
        idNhanVien: $('#IDNhanVien').val() || 0,
        doanhNghiep: window.doanhNghiep || null
    };

    $.ajax({
        url: '/bang_ke_bien_lai_hoan_ung/export/excel',
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
            a.download = `BangKeBienLaiHoanUng_${requestData.fromDate || 'all'}_den_${requestData.toDate || 'now'}.xlsx`;
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
    const idNhanVien = $('#IDNhanVien').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idNhanVien)
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
        idNhanVien: $('#IDNhanVien').val() || 0,
        doanhNghiep: window.doanhNghiep || null
    };

    fetch("/bang_ke_bien_lai_hoan_ung/export/pdf", {
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
            a.download = `BangKeBienLaiHoanUng_${requestData.fromDate || 'all'}_den_${requestData.toDate || 'now'}.pdf`;
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
    const idNhanVien = $('#IDNhanVien').val() || 0;

    if (!window.filteredData || (totalRecords && window.filteredData.length < totalRecords)) {
        fetchAllFilteredData(tu, den, idNhanVien)
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

    let hours = date.getHours();
    const minutes = String(date.getMinutes()).padStart(2, '0');

    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // 0 giờ thì chuyển thành 12

    const hh = String(hours).padStart(2, '0');

    return `${day}-${month}-${year} ${hh}:${minutes} ${ampm}`;
}

function formatCurrency(value) {
    return (value || 0.00).toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}


// ==================== CẬP NHẬT BẢNG ====================
function updateTable(response) {
    const tbody = $('.container_BangKeThu.right tbody');
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
        data.forEach((item, index) => {
            const stt = (currentPage - 1) * pageSize + index + 1;
            const row = `
                <tr>
                    <td class="text-nowrap text-center">${stt}</td>
                    <td class="text-nowrap text-center">${formatDate(item.ngayThu || item.NgayThu)}</td>
                    <td class="text-nowrap text-center">${item.maYTe || item.MaYTe || ''}</td>
                    <td class="text-nowrap text-center">${item.soBA || item.SoBA || ''}</td>
                    <td class="text-nowrap text-center">${item.maDot || item.MaDot || ''}</td>
                    <td class="text-nowrap text-start">${item.hoTenBenhNhan || item.HoTenBenhNhan || 'Không rõ'}</td>
                    <td class="text-nowrap text-center">${item.soBLHoanUng || item.SoBLHoanUng || 'Không rõ'}</td>
                    <td class="text-nowrap text-center">${item.soBLTamUng || item.SoBLTamUng || 'Không rõ'}</td>
                    <td class="text-nowrap text-end">${formatCurrency(item.giaTriHoanUng || item.GiaTriHoanUng || 0)}</td>
                    <td class="text-nowrap text-end">${formatCurrency(item.huy || item.Huy || 0)}</td>
                    <td class="text-nowrap text-end">${formatCurrency(item.hoanTra || item.HoanTra || 0)}</td>
                    <td class="text-nowrap text-center">${item.httt || item.HTTT || 'Không rõ'}</td>
                </tr>
            `;
            tbody.append(row);
        });
    } else {
        tbody.append('<tr><td colspan="12" class="text-center">Không có dữ liệu</td></tr>');
    }
}

// ==================== RÀNG BUỘC ĐIỀU KIỆN CHỌN NGÀY ====================
$(document).ready(function () {
    $('#datepicker').on('changeDate', function (e) {
        let startDate = $('#ngayTuNgay').datepicker('getDate');
        let endDate = $('#ngayDenNgay').datepicker('getDate');

        if (endDate && startDate > endDate) {
            $('#ngayDenNgay').datepicker('setDate', startDate);
        }
    });

    $('#datepicker2').on('changeDate', function (e) {
        let startDate = $('#ngayTuNgay').datepicker('getDate');
        let endDate = $('#ngayDenNgay').datepicker('getDate');

        if (startDate && endDate < startDate) {
            $('#ngayTuNgay').datepicker('setDate', endDate);
        }
    });
});

// ==================== KHI TẢI TRANG ====================
$(document).ready(function () {
    initDateInputFormatting();
    initDatePicker();
});

// ==================== LOAD COMBOBOX ====================
document.addEventListener("DOMContentLoaded", () => {
    // Initialize second combobox
    initAutocomplete({
        inputId: 'comboBox2',
        dropdownId: 'dropdownList2',
        hiddenIdId: 'IDNhanVien',
        data: provincesDataNhanVien,
        getName: item => item.Ten || '',
        getId: item => item.Id,
        getAbbr: item => item.Viettat || '',
        filterPredicate: (item, normalizedFilter) =>
            removeAccents((item.Ten || '').toLowerCase()).includes(normalizedFilter) ||
            removeAccents((item.Viettat || "").toLowerCase()).startsWith(normalizedFilter)
    });
});

// ==================== THÔNG BÁO ====================
$(document).ready(function () {
    // Chỉ hiển thị toastr nếu có tham số cụ thể trong URL
    if (window.location.search.includes('showToast=true')) {
        var successMessage = '@Html.Raw(TempData["SuccessMessage"] as string)';
        if (successMessage) {
            toastr.success(decodeHTMLEntities(successMessage));
        }

        var errorMessage = '@Html.Raw(TempData["ErrorMessage"] as string)';
        if (errorMessage) {
            toastr.error(decodeHTMLEntities(errorMessage));
        }

        var warningMessage = '@Html.Raw(TempData["WarningMessage"] as string)';
        if (warningMessage) {
            toastr.warning(decodeHTMLEntities(warningMessage));
        }
    }

    function decodeHTMLEntities(text) {
        var textArea = document.createElement('textarea');
        textArea.innerHTML = text;
        return textArea.value;
    }
});

