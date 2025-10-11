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
        url: '/bang_ke_thu_ngoai_tru/filter',
        type: 'POST',
        data: payload,
        success: function (response) {
            console.log(response);
            console.log(payload);
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

function fetchAllFilteredData(tuNgay, denNgay, idHTTT, idNhanVien, idLoai) {
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
        idHTTT: $('.tomselect-httt').val() || 0,
        idNhanVien: $('.tomselect-nhanVien').val() || 0,
        idLoai: $('.tomselect-loai').val() || 0,
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
        doanhNghiep: window.doanhNghiep || null
    };

    btnElem.disabled = true;
    btnElem.innerHTML = '<i class="bi bi-hourglass-split"></i> Đang xử lý...';


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

    if (data.length > 0) {
        const groupedData = {};
        data.forEach(item => {
            const nhanVien = item.tenNhanVien || item.TenNhanVien || "Không rõ";
            if (!groupedData[nhanVien]) {
                groupedData[nhanVien] = [];
            }
            groupedData[nhanVien].push(item);
        });

        let globalIndex = 0;

        Object.keys(groupedData).forEach(nhanVien => { 
            let tongTienNV = 0;
            let tongHuyNV = 0;
            let tongHoanTraNV = 0;
            const itemsForNhanVien = groupedData[nhanVien] || [];

            const groupedByQuyenSo = {};
            itemsForNhanVien
                .sort((a, b) => {
                    const qsA = (a.quyenSo || a.QuyenSo || "").toString();
                    const qsB = (b.quyenSo || b.QuyenSo || "").toString();
                    return qsA.localeCompare(qsB, 'vi', { sensitivity: 'base' });
                })
                .forEach(item => {
                    const quyenSo = item.quyenSo || item.QuyenSo || "Không rõ";
                    if (!groupedByQuyenSo[quyenSo]) groupedByQuyenSo[quyenSo] = [];
                    groupedByQuyenSo[quyenSo].push(item);
                });

            const nvRow = `
            <tr class="fw-bold">
                <td colspan="10" class="text-start">Nhân viên: ${nhanVien}</td>
            </tr>`;
            tbody.append(nvRow);

            Object.keys(groupedByQuyenSo).forEach(quyenSo => {
                let tongThuPhiQS = 0;
                let tongHuyQS = 0;
                let tongHoanTraQS = 0;

                const quyenSoList = groupedByQuyenSo[quyenSo];

                quyenSoList.forEach(item => {
                    tongThuPhiQS += Number(item.soTien || item.SoTien || 0);
                    tongHuyQS += Number(item.huy || item.Huy || 0);
                    tongHoanTraQS += Number(item.hoan || item.Hoan || 0);
                });

                const maNV = quyenSoList[0].maNhanVien || quyenSoList[0].MaNhanVien || "";
                const qsHeaderRow = `
                <tr class="fw-bold table-light">
                    <td colspan="3" class="text-start" style="padding-left:40px; border-right: none;">${maNV} - ${quyenSo}</td>
                    <td colspan="4" class="text-start khongapdung" style="border-left: none"></td>
                    <td class="text-end khongapdung">${formatCurrency(tongHuyQS)}</td>
                    <td class="text-end">${formatCurrency(tongHoanTraQS)}</td>
                    <td class="text-end">${formatCurrency(tongThuPhiQS)}</td>
                </tr>`;
                tbody.append(qsHeaderRow);

                quyenSoList.forEach(item => {
                    globalIndex++;
                    const stt = (currentPage - 1) * pageSize + globalIndex;
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
                tongTienNV += tongThuPhiQS;
                tongHuyNV += tongHuyQS;
                tongHoanTraNV += tongHoanTraQS;
            });

            const totalRowNV = `
            <tr class="fw-bold table-secondary">
                <td colspan="3" class="text-end">Tổng nhân viên:</td>
                <td colspan="4" class="text-end khongapdung"></td>
                <td class="text-end khongapdung">${formatCurrency(tongHuyNV)}</td>
                <td class="text-end">${formatCurrency(tongHoanTraNV)}</td>
                <td class="text-end">${formatCurrency(tongTienNV)}</td>
            </tr>`;
            tbody.append(totalRowNV);
        });

        const totalRowPage = `
        <tr class="fw-bold">
            <td colspan="3" class="text-end">Tổng tiền phải nộp:</td>
            <td colspan="4" class="text-start khongapdung">${formatCurrency(response.tongSoTien - response.tongHuy - response.tongHoan)}</td>
            <td class="text-end khongapdung">${formatCurrency(response.tongHuy || response.TongHuy)}</td>
            <td class="text-end">${formatCurrency(response.tongHoan || response.TongHoan)}</td>
            <td class="text-end">${formatCurrency(response.tongSoTien || response.TongSoTien)}</td>
        </tr>`;
        tbody.append(totalRowPage);

    } else {
        tbody.append('<tr><td colspan="10" class="text-center">Không có dữ liệu</td></tr>');
    }
}

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

const selectorBatDauKetThuc = '#ngayTuNgay, #ngayDenNgay';

$(selectorBatDauKetThuc).datetimepicker({
    format: 'DD-MM-YYYY HH:mm:ss',
    locale: 'vi',
    useCurrent: false,
    showTodayButton: true,
    showClear: true,
    showClose: true,
    calendarWeeks: false,
    tooltips: {
        today: 'Chuyển đến hôm nay',
        clear: 'Xóa lựa chọn',
        close: 'Đóng',
        selectMonth: 'Chọn tháng',
        prevMonth: 'Tháng trước',
        nextMonth: 'Tháng sau',
        selectYear: 'Chọn năm',
        prevYear: 'Năm trước',
        nextYear: 'Năm sau',
        selectTime: 'Chọn giờ'
    },
    icons: {
        time: 'ti ti-clock',
        date: 'ti ti-calendar-event',
        up: 'ti ti-chevron-up',
        down: 'ti ti-chevron-down',
        previous: 'ti ti-chevron-left',
        next: 'ti ti-chevron-right',
        today: '',
        clear: 'ti ti-trash',
        close: 'ti ti-x'
    }
})
    .on('dp.show', function () {
        const $widget = $('.bootstrap-datetimepicker-widget:last');
        const $input = $(this);
        const offset = $input.offset();
        const inputHeight = $input.outerHeight();
        const widgetHeight = $widget.outerHeight();
        const widgetWidth = $widget.outerWidth();
        const winWidth = $(window).width();
        const winHeight = $(window).height();
        const scrollTop = $(window).scrollTop();

        $widget.css({
            'transform': 'scale(0.85)',
            'transform-origin': 'top center',
        });

        const scaledHeight = widgetHeight * 0.85;
        const scaledWidth = widgetWidth * 0.85;

        let left = offset.left + $input.outerWidth() + 10 - 180;
        let top = offset.top + 40;

        if (left + scaledWidth > winWidth - 10) {
            left = offset.left - scaledWidth - 10;
        }

        if (top < scrollTop + 10) {
            top = offset.top + inputHeight + 10;
        }

        if (left < 10) left = 10;
        if (left + scaledWidth > winWidth - 10) {
            left = winWidth - scaledWidth - 10;
        }
        if (top < scrollTop + 10) top = scrollTop + 10;
        if (top + scaledHeight > scrollTop + winHeight - 10) {
            top = scrollTop + winHeight - scaledHeight - 10;
        }

        $widget.appendTo('body').css({
            position: 'absolute',
            top: top,
            left: left,
            zIndex: 999999
        }).addClass('active-popup');
    })
    .on('dp.hide', function () {
        $('.bootstrap-datetimepicker-widget')
            .removeClass('active-popup')
            .css('transform', '');
    });
    Inputmask({
        alias: "datetime",
        inputFormat: "dd-mm-yyyy HH:MM:ss",
        placeholder: "dd-mm-yyyy hh:mm:ss",
        clearIncomplete: true,
        showMaskOnHover: false,
        showMaskOnFocus: true
    }).mask(selectorBatDauKetThuc);

    const style = document.createElement('style');
    style.textContent = `
                #ngayTuNgay,
                #ngayDenNgay{
                    background-color: #f8f9fa;
                    border: 1px solid #ced4da;
                }

                .bootstrap-datetimepicker-widget {
                    border: 1px solid #ccc;
                    border-radius: 8px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                }
            `;
document.head.appendChild(style);