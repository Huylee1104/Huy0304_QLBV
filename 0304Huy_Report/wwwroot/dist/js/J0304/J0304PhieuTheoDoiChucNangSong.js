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

    //input.addEventListener('blur', () => {
    //    setTimeout(() => {
    //        if(!isMouseDownOnDropdown) {
    //            if (hiddenId.value === "" && input.value.trim() !== "") {
    //                input.value = "";
    //                hiddenId.value = 0;
    //                toastr.error("Vui lòng chọn bệnh nhân hợp lệ");
    //            }
    //            if (hiddenId.value === "0") {
    //                toastr.error("Vui lòng chọn bệnh nhân");
    //            }
    //        }
    //        isMouseDownOnDropdown = false;
    //        dropdown.style.display = "none";
    //    }, 100);
    //});

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
    let idBenhNhan = $('#IDBenhNhan').val();
    if (!isPagination) {
        firstLoad = true;
    }

    $('#loadingSpinner').show();
    $('.table-wrapper').css('opacity', '0.5');

    let payload = {
        IdChiNhanh: _idcn,
        idVaoVien: 9,
        page: currentPage,
        pageSize: pageSize
    }
    $.ajax({
        url: '/phieu_theo_doi_chuc_nang_song/filter',
        type: 'POST',
        data: payload,
        success: function (response) {
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
            url: '/phieu_theo_doi_chuc_nang_song/filter',
            type: 'POST',
            data: payload,
            success: function (resp) { resolve(resp); },
            error: function (xhr, st, err) { reject(err || st || xhr); }
        });
    });
}

function fetchAllFilteredData(idBenhNhan) {
    return new Promise((resolve, reject) => {
        const basePayload = {
            IdChiNhanh: _idcn || 0,
            idBenhNhan: idBenhNhan || 0,
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
                    IdChiNhanh: _idcn,
                    idBenhNhan: idBenhNhan,
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
    //const idBenhNhan = $('#IDBenhNhan').val();
    //if (idBenhNhan === "0" || idBenhNhan === "" || idBenhNhan == null) {
    //    toastr.error("Vui lòng chọn bệnh nhân");
    //    return false;
    //}
    //if (!window.filteredData || window.filteredData.length === 0) {
    //    toastr.error("Không có dữ liệu để xuất");
    //    return false;
    //}
    return true;
}

// ==================== XUẤT PDF ====================
function doExportPdf(finalData, btnElem) {
    const requestData = {
        data: finalData,
        idBenhNhan: $('#IDBenhNhan').val(),
        doanhNghiep: window.doanhNghiep || null
    };
    console.log(requestData, );
    console.log("data json: ", JSON.stringify(requestData));
    fetch("/phieu_theo_doi_chuc_nang_song/export/pdf", {
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

    const idBenhNhan = $('#IDBenhNhan').val();
    
    if (!window.filteredData) {
        (idBenhNhan)
            fetchAllFilteredData.then(allData => {
                window.filteredData = allData;
                doExportPdf(allData, btn);
            })
            .catch(err => {
                btn.innerHTML = '<i class="bi bi-file-earmark-pdf"></i> Xuất PDF';
                btn.disabled = false;
            });
    } else {
        console.log(window.filteredData[0]);
        doExportPdf(window.filteredData[0], btn);
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
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');

    return `${day}-${month}-${year} ${hours}:${minutes}`;
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
    if (response.data.thongTinBenhNhan) {
        const info = response.data.thongTinBenhNhan;
        const rowInfo = `
        <tr class="table-info">
            <td colspan="7" class="text-start">
                <strong>Họ tên:</strong> ${info.tenBenhNhan || info.TenBenhNhan || ''} <br>
                <strong>Mã bệnh nhân:</strong> ${info.maBenhNhan || info.MaBenhNhan || ''} <br>
                <strong>Khoa:</strong> ${info.tenKhoa || info.TenKhoa || ''} <br>
                <strong>Tên phòng:</strong> ${info.tenPhong || info.TenPhong || ''} <br>
                <strong>Tên buồng:</strong> ${info.tenGiuong || info.TenGiuong || ''} <br>
                <strong>Tuổi:</strong> ${info.ngaySinh || info.NgaySinh || ''} <br>
                <strong>Địa chỉ:</strong> ${info.diaChi || info.DiaChi || ''} <br>
                <strong>Giới tính:</strong> ${info.gioiTinh || info.GioiTinh || ''} <br>
                <strong>Chẩn đoán:</strong> ${info.chanDoan || info.ChanDoan || ''} <br>
                <strong>Mã vào viện:</strong> ${info.maVaoVien || info.MaVaovien || ''}
            </td>
        </tr>
    `;
        tbody.append(rowInfo);
    }

    // Render danh sách sinh hiệu
    if (response.data.sinhHieus && response.data.sinhHieus.length > 0) {
        response.data.sinhHieus.forEach((item, index) => {
            const row = `
                <tr>
                    <td class="text-center">${index + 1}</td>
                    <td class="text-center">${formatDate(item.ngayKhaoSat || item.NgayKhaoSat)}</td>
                    <td class="text-center">${item.mach || item.Mach ||''}</td>
                    <td class="text-center">${item.nhietDo || item.NhietDo || ''}</td>
                    <td class="text-center">${item.huyetAp || item.HuyetAp || ''}</td>
                    <td class="text-center">${item.canNang || item.CanNang || ''}</td>
                    <td class="text-center">${item.nhipTho || item.NhipTho || ''}</td>
                </tr>
            `;
            tbody.append(row);
        });
    } else {
        tbody.append('<tr><td colspan="7" class="text-center">Không có dữ liệu sinh hiệu</td></tr>');
    }
}

// ==================== LOAD COMBOBOX ====================
document.addEventListener("DOMContentLoaded", () => {
    const dataBenhNhan = convertData(
        provincesDataBenhNhan,
        item => item.TenBN || '',
        item => item.ID
    );

    initAutocomplete({
        inputId: 'comboBox',
        dropdownId: 'dropdownList',
        hiddenIdId: 'IDBenhNhan',
        data: dataBenhNhan,
        getName: item => item.ten || '',
        getId: item => item.id,
        getAbbr: item => item.viettat || '',
        filterPredicate: (item, normalizedFilter) =>
            removeAccents((item.ten || '').toLowerCase()).includes(normalizedFilter) ||
            removeAccents((item.viettat || '').toLowerCase()).startsWith(normalizedFilter)
    });
});

