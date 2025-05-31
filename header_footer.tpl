{% extends "html_table.tpl" %}

{% block before_table %}
{% if true %}
<script src="https://unpkg.com/pagedjs/dist/paged.polyfill.js"></script>
{% endif %}
<style>
@page {
  @top-left {
    content: running(topLeft);
  }
  @top-center {
    content: running(topCenter);
  }
  @top-right {
    content: running(topRight);
  }
  @bottom-left {
    {% if pagenum_bottom_left is sameas true %}
    content: "Page " counter(page) " of " counter(pages);
    {{ style_footer_left| default("") }}
    {% else %}
    content: running(bottomLeft);
    {% endif %}
  }
  @bottom-center {
    content: running(bottomCenter);
  }
  @bottom-right {
    {% if pagenum_bottom_right is sameas true %}
    content: "Page " counter(page) " of " counter(pages);
    {{ style_footer_right | default("") }}
    {% else %}
    content: running(bottomRight);
    {% endif %}
  }
  .hidden-print {
  }

  .top-left {
    position: running(topLeft);
  }
  .top-center {
    position: running(topCenter);
  }
  .top-right {
    position: running(topRight);
  }
  .bottom-left {
    position: running(bottomLeft);
  }
  .bottom-center {
    position: running(bottomCenter);
  }
  .bottom-right {
    position: running(bottomRight);
  }
}

table {
    --pagedjs-repeat-header: all;
}
</style>
{% endblock before_table %}

{% block before_head_rows %}
<tr class="hidden-print">
    <th colspan=42>
        <div id="header_{{uuid}}" style="display:grid;grid-auto-flow:column;grid-auto-columns:1fr">
            <div class="top-left" id="header_left_{{uuid}}" style="text-align:left;{{ style_header_left  | default('') }}">{{ content_header_left | default("") }}</div>
            <div class="top-center" id="header_center_{{uuid}}" style="text-align:center;{{ style_header_center  | default('') }}">{{ content_header_center | default("") }}</div>
            <div class="top-right" id="header_right_{{uuid}}" style="text-align:right;{{ style_header_right  | default('') }}">{{ content_header_right | default("") }}</div>
        </div>
    </th>
</tr>
{{ super() }}
{% endblock before_head_rows %}

{% block tbody %}
{{ super() }}
{% block tfoot %}
<tfoot class="hidden-print">
    <tr>
        <th colspan=42>
            <div id="footer_{{uuid}}" style="display:grid;grid-auto-flow:column;grid-auto-columns:1fr">
                <div class="bottom-left" id="footer_left_{{uuid}}" style="text-align:left;{{ style_footer_left | default('') }}">{{ content_footer_left | default("") }}</div>
                <div class="bottom-center" id="footer_center_{{uuid}}" style="text-align:center;{{ style_footer_center | default('') }}">{{ content_footer_center | default("") }}</div>
                <div class="bottom-right" id="footer_right_{{uuid}}" style="text-align:right;{{ style_footer_right | default('') }}">{{ content_footer_right | default("") }}</div>
            </div>
        </th>
    </tr>
</tfoot>
{% endblock tfoot %}
{% endblock tbody %}
