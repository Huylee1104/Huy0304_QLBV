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
                toastr.error("Không thể tải dữ liệu");
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
                    <td style="width: 5%; text-align:center;">${stt}</td>
                    <td style="width: 10%; text-align:center;">${item.maYTe || item.MaYTe || ''}</td>
                    <td style="width: 20%; text-align:left;">${item.HoVaTen || item.hoVaTen || 'Không rõ'}</td>
                    <td style="width: 10%; text-align:center;">${item.quyenSo || item.QuyenSo || 'Không rõ'}</td>
                    <td style="width: 10%; text-align:center;">${item.soBienLai || item.SoBienLai || 'Không rõ'}</td>
                    <td style="width: 5%; text-align:center;">${item.loai || item.Loai || 'Không rõ'}</td>
                    <td style="width: 10%; text-align:center;">${formatDate(item.ngayThu || item.NgayThu)}</td>
                    <td style="width: 10%; text-align:right;">${formatCurrency(item.huy || item.Huy)}</td>
                    <td style="width: 10%; text-align:right;">${formatCurrency(item.hoan || item.Hoan)}</td>
                    <td style="width: 10%; text-align:right;">${formatCurrency(item.soTien || item.SoTien)}</td>
                </tr>
            `;
            tbody.append(row);
        });
    } else {
        tbody.append('<tr><td colspan="10" class="text-center">Không có dữ liệu</td></tr>');
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

// ==================== CHỌN NGÀY THEO NĂM, QUÝ, THÁNG, NGÀY ====================
$('#selectGiaiDoan').change(function () {
    const selectedValue = $(this).val();
    const container = $('#selectContainer');
    container.empty();

    if (selectedValue === 'Nam' || selectedValue === 'Ngay') {
        container.css('justify-content', 'flex-start');
    } else if (selectedValue === 'Quy' || selectedValue === 'Thang') {
        container.css('justify-content', 'space-around');
    }

    const currentYear = new Date().getFullYear();
    const currentMonth = new Date().getMonth() + 1;
    const currentQuy = Math.ceil(currentMonth / 3);

    // ================== FUNCTION TẠO DROPDOWN ==================
    function createDropdownInput(id, label, values, defaultValue, onSelect, length = 10) {
        const html = `
            <div data-dropdown-wrapper style="width: 46%; position: relative;">
                <label class="form-label">${label}</label>
                <input type="number" class="form-control" id="${id}" value="${defaultValue}" oninput="if(this.value.length > ${length}) this.value = this.value.slice(0, ${length});"  autocomplete="off">
                <div id="${id}Dropdown"
                    style="display:none; position:absolute; top:100%; left:0; width:100%;
                    max-height:200px; overflow-y:auto; z-index:9999; background:white;
                    border:1px solid rgba(0,0,0,.15); border-radius:4px;
                    box-shadow:0 6px 12px rgba(0,0,0,.175);">
                </div>
            </div>
        `;
        container.append(html);

        const $input = $('#' + id);
        const $dropdown = $('#' + id + 'Dropdown');
        let currentHighlightIndex = -1;

        function highlightCurrentItem() {
            const items = $dropdown.find('.dropdown-item');
            items.removeClass('active bg-primary text-white');
            if (currentHighlightIndex >= 0 && currentHighlightIndex < items.length) {
                items.eq(currentHighlightIndex).addClass('active bg-primary text-white');
                const item = items.eq(currentHighlightIndex)[0];
                if (item) item.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            }
        }

        // Trong hàm renderList(), sửa phần kiểm tra giá trị như sau:
        function renderList(filter = '') {
            $dropdown.empty();
            currentHighlightIndex = -1;

            const typedVal = parseInt($input.val(), 10);
            const typedIsAllowed = Number.isFinite(typedVal) && (values.includes(typedVal) || id === 'yearInput');

            // Xác định giá trị hiện tại để highlight
            let highlightVal = typedVal;
            if ((id === 'quyInput' || id === 'thangInput') &&
                (!Number.isFinite(typedVal) ||
                    (id === 'quyInput' && (typedVal < 1 || typedVal > 4)) ||
                    (id === 'thangInput' && (typedVal < 1 || typedVal > 12)))) {

                // Lấy giá trị hiện tại để highlight nhưng không thay đổi input
                const now = new Date();
                if (id === 'quyInput') {
                    highlightVal = Math.ceil((now.getMonth() + 1) / 3);
                } else {
                    highlightVal = now.getMonth() + 1;
                }
            }

            let filteredValues = values.filter(v => !filter || v.toString().includes(filter));
            if (filteredValues.length === 0 && id === 'yearInput') {
                if (Number.isFinite(typedVal)) {
                    filteredValues = [typedVal];
                } else {
                    filteredValues = values.slice();
                }
            } else if (filteredValues.length === 0) {
                filteredValues = values.slice();
            }

            filteredValues.forEach((val, index) => {
                // Sử dụng highlightVal thay vì typedVal để xác định isSelected
                const isSelected = Number.isFinite(highlightVal) && val === highlightVal;
                const item = $(` 
            <a href="#" class="dropdown-item ${isSelected ? 'active bg-primary text-white' : ''}"
               data-val="${val}" data-index="${index}"
               style="padding:8px 16px; display:block; text-decoration:none; color:#333; cursor:pointer;">
               ${val}
            </a>
        `);
                item.on('click', function (e) {
                    e.preventDefault();
                    selectItem(val);
                });
                item.on('mouseenter', function () {
                    currentHighlightIndex = index;
                    highlightCurrentItem();
                });
                $dropdown.append(item);
                if (isSelected) currentHighlightIndex = index;
            });

            const items = $dropdown.find('.dropdown-item');
            if (currentHighlightIndex === -1 && items.length) {
                currentHighlightIndex = 0;
            }
            highlightCurrentItem();
        }

        function selectItem(val) {
            $input.val(val);
            $dropdown.hide();
            if (onSelect) onSelect(val);
        }

        $input.on('focus click', function () {
            renderList();
            $dropdown.show();
        });

        $input.on('input', function () {
            renderList($(this).val());
            $dropdown.show();
        });

        $input.on('keydown', function (e) {
            const items = $dropdown.find('.dropdown-item');
            if (!items.length) return;

            const key = e.key;
            const isUp = key === 'ArrowUp';
            const isDown = key === 'ArrowDown';
            const isEnter = key === 'Enter';
            const isEscape = key === 'Escape';
            const isTab = key === 'Tab';

            if (isUp || isDown || isEnter || isEscape || isTab) e.preventDefault();

            if (isUp) {
                currentHighlightIndex = (currentHighlightIndex <= 0) ? items.length - 1 : currentHighlightIndex - 1;
                highlightCurrentItem();
                return;
            }

            if (isDown) {
                currentHighlightIndex = (currentHighlightIndex >= items.length - 1) ? 0 : currentHighlightIndex + 1;
                highlightCurrentItem();
                return;
            }

            if (isEnter && currentHighlightIndex >= 0) {
                const val = parseInt(items.eq(currentHighlightIndex).data('val'), 10);
                selectItem(val);
                return;
            }

            if (isEscape) {
                $dropdown.hide();
                return;
            }

            if (isTab) {
                if (currentHighlightIndex >= 0) {
                    const val = parseInt(items.eq(currentHighlightIndex).data('val'), 10);
                    selectItem(val);
                }
                return;
            }
        });
        $input.on('keypress', function (e) {
            const invalidChars = ['e', 'E', '+', '-', '.', ','];
            if (invalidChars.includes(e.key)) {
                e.preventDefault();
            }
        });
        $(document).off('click.dropdown-' + id).on('click.dropdown-' + id, function (e) {
            if (!$(e.target).closest('[data-dropdown-wrapper]').length) {
                $dropdown.hide();
            }
        });
    }

    // ================== FORMAT DATE ==================
    function formatDate(date) {
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}-${month}-${year}`;
    }

    function getMonthDateRange(year, month) {
        const startDate = new Date(year, month - 1, 1);
        const endDate = new Date(year, month, 0);
        return { start: startDate, end: endDate };
    }

    function highlightYearInDropdown(year) {
        $('#yearInputDropdown').find('.dropdown-item').removeClass('active bg-primary text-white');
        const yearItem = $('#yearInputDropdown').find(`[data-val="${year}"]`);
        if (yearItem.length) {
            yearItem.addClass('active bg-primary text-white');
            yearItem[0].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
    }

    // ================== UPDATE DATE RANGE ==================
    function updateDates() {
        let yearRaw = parseInt($('#yearInput').val(), 10);
        let year = Number.isFinite(yearRaw) ? yearRaw : currentYear;

        // Chỉ kiểm tra năm không âm
        if (year < 0 || year > currentYear) {
            year = currentYear;
            $('#yearInput').val(currentYear);
            highlightYearInDropdown(currentYear);
        }

        if (selectedValue === 'Nam') {
            $('#ngayTuNgay').val(`01-01-${year}`);
            $('#ngayDenNgay').val(`31-12-${year}`);
        }
        else if (selectedValue === 'Quy') {
            let quy = parseInt($('#quyInput').val(), 10);
            if (!Number.isFinite(quy)) quy = currentQuy;
            if (quy < 1) quy = 1;
            if (quy > 4) quy = 4;
            $('#quyInput').val(quy);

            const startMonth = (quy - 1) * 3 + 1;
            const endMonth = startMonth + 2;
            $('#ngayTuNgay').val(formatDate(new Date(year, startMonth - 1, 1)));
            $('#ngayDenNgay').val(formatDate(new Date(year, endMonth, 0)));
        }
        else if (selectedValue === 'Thang') {
            let month = parseInt($('#thangInput').val(), 10);
            if (!Number.isFinite(month)) month = currentMonth;
            if (month < 1) month = 1;
            if (month > 12) month = 12;
            $('#thangInput').val(month);

            const { start, end } = getMonthDateRange(year, month);
            $('#ngayTuNgay').val(formatDate(start));
            $('#ngayDenNgay').val(formatDate(end));
        }
        else if (selectedValue === 'Ngay') {
            const today = new Date(Date.now());
            const todayStr = formatDate(today);
            $('#ngayTuNgay').val(todayStr);
            $('#ngayDenNgay').val(todayStr);
        }

        if (selectedValue === 'Nam' || selectedValue === 'Quy' || selectedValue === 'Thang') {
            $('#ngayTuNgay, #ngayDenNgay').prop('disabled', true);
        } else {
            $('#ngayTuNgay, #ngayDenNgay').prop('disabled', false);
        }

        $('#ngayTuNgay').datepicker('setDate', $('#ngayTuNgay').val());
        $('#ngayDenNgay').datepicker('setDate', $('#ngayDenNgay').val());
    }

    const startYear = 2000;
    const yearOptions = Array.from({ length: currentYear - startYear + 1 }, (_, i) => startYear + i);
    createDropdownInput('yearInput', 'Năm', yearOptions, currentYear, updateDates, 4);
    $(document)
        .off('blur', '#yearInput')
        .on('blur', '#yearInput', function () {
            let val = parseInt($(this).val(), 10);
            if (!Number.isFinite(val) || val > currentYear || val < 0) val = currentYear;
            $(this).val(val);

            $('#quyInputDropdown').find('.dropdown-item').removeClass('active bg-primary text-white');
            $('#quyInputDropdown').find(`[data-val="${val}"]`).addClass('active bg-primary text-white');

            updateDates();
        });

    // ================== QUÝ ==================
    if (selectedValue === 'Quy') {
        createDropdownInput('quyInput', 'Quý', [1, 2, 3, 4], currentQuy, updateDates, 1);

        $(document)
            .off('blur', '#quyInput')
            .on('blur', '#quyInput', function () {
                let val = parseInt($(this).val(), 10);
                if (!Number.isFinite(val) || val < 1 || val > 4) val = currentQuy;
                $(this).val(val);

                $('#quyInputDropdown').find('.dropdown-item').removeClass('active bg-primary text-white');
                $('#quyInputDropdown').find(`[data-val="${val}"]`).addClass('active bg-primary text-white');

                updateDates();
            });
    }

    // ================== THÁNG ==================
    else if (selectedValue === 'Thang') {
        createDropdownInput('thangInput', 'Tháng', Array.from({ length: 12 }, (_, i) => i + 1), currentMonth, updateDates, 2);

        $(document)
            .off('blur', '#thangInput')
            .on('blur', '#thangInput', function () {
                let val = parseInt($(this).val(), 10);
                if (!Number.isFinite(val) || val < 1 || val > 12) val = currentMonth;
                $(this).val(val);

                $('#thangInputDropdown').find('.dropdown-item').removeClass('active bg-primary text-white');
                $('#thangInputDropdown').find(`[data-val="${val}"]`).addClass('active bg-primary text-white');

                updateDates();
            });
    }

    else if (selectedValue === 'Ngay') {
        container.empty();
    }

    updateDates();
});

// ==================== KHI TẢI TRANG ====================
$(document).ready(function () {
    initDateInputFormatting();
    initDatePicker();
});

// ==================== LOAD COMBOBOX ====================
document.addEventListener("DOMContentLoaded", () => {
    // Initialize first combobox
    initAutocomplete({
        inputId: 'comboBox',
        dropdownId: 'dropdownList',
        hiddenIdId: 'IDHTTT',
        data: provincesDataHTTT,
        getName: item => item.ten || '',
        getId: item => item.id,
        filterPredicate: (item, normalizedFilter) =>
            removeAccents((item.ten || '').toLowerCase()).includes(normalizedFilter)
    });

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

